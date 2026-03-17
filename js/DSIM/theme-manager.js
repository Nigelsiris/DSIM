// Theme manager for DSIM
class ThemeManager {
    constructor() {
        this.THEME_KEY = 'dsim_theme';
        this.initTheme();
    }

    initTheme() {
        const savedTheme = localStorage.getItem(this.THEME_KEY) || 'light';
        this.setTheme(savedTheme);
        this.addThemeToggle();
    }

    setTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem(this.THEME_KEY, theme);
    }

    toggleTheme() {
        const currentTheme = document.documentElement.getAttribute('data-theme');
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        this.setTheme(newTheme);
    }

    addThemeToggle() {
        const toggle = document.createElement('button');
        toggle.id = 'theme-toggle';
        toggle.className = 'theme-toggle';
        toggle.innerHTML = '🌓';
        toggle.onclick = () => this.toggleTheme();
        
        document.body.appendChild(toggle);
    }
}

// Initialize theme manager
window.themeManager = new ThemeManager();