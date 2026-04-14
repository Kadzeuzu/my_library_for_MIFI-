(function() {
    // 1. КОНФИГУРАЦИЯ ТЕМ
    const themes = [
        { name: 'default', icon: '🌙', label: 'Тёмная', themeColor: '#0a0a0a' },
        { name: 'warm', icon: '☕', label: 'Тёплая', themeColor: '#fdf6e3' },
        { name: 'light', icon: '☀️', label: 'Светлая', themeColor: '#f8f9fa' }
    ];

    let currentThemeIndex = 0;
    const savedTheme = localStorage.getItem('schedule-theme');
    if (savedTheme) {
        const idx = themes.findIndex(t => t.name === savedTheme);
        if (idx !== -1) currentThemeIndex = idx;
    }
    applyTheme(themes[currentThemeIndex].name);

    function applyTheme(themeName) {
        // Установка data-theme
        if (themeName === 'default') {
            document.documentElement.removeAttribute('data-theme');
        } else {
            document.documentElement.setAttribute('data-theme', themeName);
        }

        // Обновление meta theme-color из массива
        const themeMeta = document.querySelector('meta[name="theme-color"]');
        if (themeMeta) {
            const theme = themes.find(t => t.name === themeName) || themes[0];
            themeMeta.setAttribute('content', theme.themeColor);
        }
    }

    function updateToggleButton(btn, index) {
        const theme = themes[index];
        btn.innerHTML = `${theme.icon} <span>${theme.label}</span>`;
        btn.title = `Сменить тему (сейчас: ${theme.label})`;
    }

    // Инициализация переключателя тем
    function initThemeToggle() {
        const toggleBtn = document.getElementById('themeToggle');
        if (!toggleBtn) return;

        const oldHandler = toggleBtn._themeHandler;
        if (oldHandler) {
            toggleBtn.removeEventListener('click', oldHandler);
        }

        const handler = () => {
            currentThemeIndex = (currentThemeIndex + 1) % themes.length;
            const themeName = themes[currentThemeIndex].name;
            applyTheme(themeName);
            updateToggleButton(toggleBtn, currentThemeIndex);
            localStorage.setItem('schedule-theme', themeName);
        };

        toggleBtn.addEventListener('click', handler);
        toggleBtn._themeHandler = handler;

        updateToggleButton(toggleBtn, currentThemeIndex);
    }

    function timeToMinutes(timeStr) {
        if (!timeStr) return 0;
        const startTime = timeStr.split('-')[0].trim();
        const parts = startTime.split(/[.:]/);
        const hours = parseInt(parts[0], 10) || 0;
        const minutes = parseInt(parts[1], 10) || 0;
        return hours * 60 + minutes;
    }

    function getLessonTypeClass(lessonName) {
        const nameLower = lessonName.toLowerCase();
        if (nameLower.includes('(лек)') || nameLower.includes('лекция')) return 'lecture';
        if (nameLower.includes('(лаб)') || nameLower.includes('лабораторная')) return 'lab';
        if (nameLower.includes('(пр)') || nameLower.includes('практика')) return 'practice';
        if (nameLower.includes('язык')) return 'practice';
        return '';
    }

    initThemeToggle();

    fetch('data.json')
        .then(response => {
            if (!response.ok) throw new Error('Ошибка загрузки');
            return response.json();
        })
        .then(data => {
            const titleEl = document.getElementById('week-title');
            if (titleEl) titleEl.innerText = data.week_title;
            
            const downloadLink = document.getElementById('download-link');
            if (downloadLink) downloadLink.href = data.original_link;
            
            const isUpper = data.week_title.toLowerCase().includes('верхняя');
            const container = document.getElementById('schedule');
            container.innerHTML = '';
            
            data.days.forEach(day => {
                const lessonsByTime = {};
                day.lessons.forEach(lesson => {
                    const time = lesson.time;
                    if (!lessonsByTime[time]) lessonsByTime[time] = [];
                    lessonsByTime[time].push(lesson);
                });
                
                const filteredLessons = [];
                
                for (const time in lessonsByTime) {
                    let group = lessonsByTime[time];
                    let candidates = [...group];
                    
                    const hasPhysicsLab = candidates.some(l => 
                        l.name.toLowerCase().includes('физика') && 
                        (l.name.toLowerCase().includes('лаб') || l.name.toLowerCase().includes('лабораторная'))
                    );
                    
                    if (hasPhysicsLab) {
                        if (isUpper) {
                            candidates = candidates.filter(l => 
                                !(l.name.toLowerCase().includes('физика') && 
                                  (l.name.toLowerCase().includes('лаб') || l.name.toLowerCase().includes('лабораторная')))
                            );
                        } else {
                            candidates = candidates.filter(l => 
                                l.name.toLowerCase().includes('физика') && 
                                (l.name.toLowerCase().includes('лаб') || l.name.toLowerCase().includes('лабораторная'))
                            );
                        }
                    }
                    
                    if (candidates.length === 0) continue;
                    
                    let selected;
                    if (candidates.length === 1) {
                        selected = candidates[0];
                    } else {
                        selected = isUpper ? candidates[0] : (candidates[1] || candidates[0]);
                    }
                    
                    filteredLessons.push(selected);
                }
                
                filteredLessons.sort((a, b) => {
                    return timeToMinutes(a.time) - timeToMinutes(b.time);
                });
                
                let dayHtml = `<div class="day-card"><h3>${day.day} (${day.date})</h3>`;
                
                filteredLessons.forEach(l => {
                    const typeClass = getLessonTypeClass(l.name);
                    
                    dayHtml += `
                        <div class="lesson ${typeClass}">
                            <div class="lesson-time">${l.time}</div>
                            <div class="lesson-info">
                                <div class="lesson-name">${l.name}</div>
                                <div class="lesson-details">
                                    ${l.teacher || '—'} | ауд. ${l.room || '—'}
                                </div>
                            </div>
                        </div>`;
                });
                
                dayHtml += `</div>`;
                if (filteredLessons.length > 0) {
                    container.innerHTML += dayHtml;
                }
            });
        })
        .catch(error => {
            console.error('Ошибка загрузки расписания:', error);
            const container = document.getElementById('schedule');
            if (container) {
                container.innerHTML = '<p style="color: var(--text); padding: 20px; text-align: center;">Ошибка загрузки расписания</p>';
            }
        });
})();