(function (window, document) {
  const THEME_KEY = 'dsim_theme';
  const DEFAULT_OPTIONS = {
    themeToggleContainer: '#topRightControls',
    loadingOverlaySelector: '.loading-overlay',
    toastSelector: '#toast',
    autoCreateContainers: true
  };

  const sanitize = (value) => (typeof value === 'string' ? value.trim() : '');

  function ensureElement(selector, create) {
    let element = document.querySelector(selector);
    if (!element && typeof create === 'function') {
      element = create();
    }
    return element;
  }

  function ensureTopControls() {
    return ensureElement(DEFAULT_OPTIONS.themeToggleContainer, () => {
      const controls = document.createElement('div');
      controls.id = DEFAULT_OPTIONS.themeToggleContainer.replace('#', '');
      controls.className = 'top-controls';
      document.body.appendChild(controls);
      return controls;
    });
  }

  const AppShell = {
    options: { ...DEFAULT_OPTIONS },
    overlayEl: null,
    toastEl: null,
    init(options = {}) {
      this.options = { ...DEFAULT_OPTIONS, ...options };
      this.overlayEl = ensureElement(this.options.loadingOverlaySelector, () => {
        if (!this.options.autoCreateContainers) return null;
        const overlay = document.createElement('div');
        overlay.className = 'loading-overlay';
        overlay.innerHTML = '<div class="loading-spinner" role="status" aria-live="polite"></div>';
        document.body.appendChild(overlay);
        return overlay;
      });
      this.toastEl = ensureElement(this.options.toastSelector, () => {
        if (!this.options.autoCreateContainers) return null;
        const toast = document.createElement('div');
        toast.id = this.options.toastSelector.replace('#', '');
        toast.className = 'toast';
        document.body.appendChild(toast);
        return toast;
      });
      this.applySavedTheme();
      this.injectThemeToggle();
      return this;
    },
    applySavedTheme() {
      const saved = localStorage.getItem(THEME_KEY);
      const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
      this.setTheme(saved || (prefersDark ? 'dark' : 'light'));
    },
    setTheme(theme) {
      document.documentElement.setAttribute('data-theme', theme);
      localStorage.setItem(THEME_KEY, theme);
    },
    toggleTheme() {
      const current = document.documentElement.getAttribute('data-theme') || 'light';
      this.setTheme(current === 'light' ? 'dark' : 'light');
    },
    injectThemeToggle() {
      const container = document.querySelector(this.options.themeToggleContainer) || (this.options.autoCreateContainers ? ensureTopControls() : null);
      if (!container || container.querySelector('.theme-toggle')) return;
      container.classList.add('top-controls');
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'theme-toggle secondary outline';
      button.id = 'themeToggleBtn';
      button.innerHTML = '<span aria-hidden="true">🌓</span><span class="sr-only">Toggle theme</span><span class="theme-toggle__label">Theme</span>';
      button.addEventListener('click', () => this.toggleTheme());
      container.appendChild(button);
    },
    showLoading() {
      if (this.overlayEl) {
        this.overlayEl.classList.add('is-active');
      }
    },
    hideLoading() {
      if (this.overlayEl) {
        this.overlayEl.classList.remove('is-active');
      }
    },
    toast(message, { variant = 'info', duration = 2500 } = {}) {
      if (!this.toastEl) return;
      this.toastEl.textContent = message;
      this.toastEl.dataset.variant = variant;
      this.toastEl.classList.add('is-visible');
      window.clearTimeout(this.toastTimer);
      this.toastTimer = window.setTimeout(() => {
        this.toastEl.classList.remove('is-visible');
      }, duration);
    },
    setBusy(element, isBusy) {
      if (!element) return;
      if (isBusy) {
        element.setAttribute('aria-busy', 'true');
        element.dataset.originalText = element.dataset.originalText || element.textContent;
      } else {
        element.removeAttribute('aria-busy');
      }
    },
    forceUppercase(selector) {
      const element = typeof selector === 'string' ? document.querySelector(selector) : selector;
      if (!element) return;
      element.addEventListener('blur', () => {
        element.value = sanitize(element.value).toUpperCase();
      });
    },
    persistFields(storageKey, fieldsConfig = {}) {
      const entries = Object.entries(fieldsConfig).map(([key, config]) => {
        if (typeof config === 'string') {
          return { key, selector: config, transform: {} };
        }
        const { selector, transform = {}, defaultValue = '' } = config;
        return { key, selector, transform, defaultValue };
      });

      const getElement = (entry) => document.querySelector(entry.selector);

      const readValues = () => {
        const data = {};
        let hasValue = false;
        entries.forEach((entry) => {
          const el = getElement(entry);
          if (!el) return;
          let value = typeof entry.transform.read === 'function' ? entry.transform.read(el) : el.value;
          value = typeof entry.transform.save === 'function' ? entry.transform.save(value) : sanitize(value);
          data[entry.key] = value;
          if (value) {
            hasValue = true;
          }
        });
        if (hasValue) {
          localStorage.setItem(storageKey, JSON.stringify(data));
        } else {
          localStorage.removeItem(storageKey);
        }
      };

      const restoreValues = () => {
        const raw = localStorage.getItem(storageKey);
        if (!raw) return;
        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch (e) {
          localStorage.removeItem(storageKey);
          return;
        }
        entries.forEach((entry) => {
          const el = getElement(entry);
          if (!el || !(entry.key in parsed)) return;
          const value = typeof entry.transform.restore === 'function' ? entry.transform.restore(parsed[entry.key]) : parsed[entry.key];
          if (typeof entry.transform.write === 'function') {
            entry.transform.write(el, value);
          } else {
            el.value = value;
          }
        });
      };

      entries.forEach((entry) => {
        const el = getElement(entry);
        if (!el) return;
        ['change', 'blur'].forEach((evt) => el.addEventListener(evt, readValues));
      });

      restoreValues();

      return {
        save: readValues,
        restore: restoreValues,
        clear(resetFields = true) {
          localStorage.removeItem(storageKey);
          if (resetFields) {
            entries.forEach((entry) => {
              const el = getElement(entry);
              if (!el) return;
              const fallback = 'defaultValue' in entry ? entry.defaultValue : '';
              if (typeof entry.transform.write === 'function') {
                entry.transform.write(el, fallback);
              } else {
                el.value = fallback;
              }
            });
          }
        }
      };
    },
    sanitize: sanitize
  };

  window.AppShell = AppShell;
})(window, document);
