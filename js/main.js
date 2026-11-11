/**
 * ALIEF SUKMA DEWANTA — CYBERPUNK 2077 TACTICAL HUD INTERACTIVE LOGIC
 * High-performance, zero-bloat vanilla JavaScript.
 */

document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initHeader();
  initReadingProgress();
  initSurabayaClock();
  initDataShards();
  initRoleLens();
  initTerminalDrawer();
  initLightbox();
  initMediaTilt();
  initContactForm();
  initMobileMenu();
  initNavDropdown();
});

/* --- 1. THEME TOGGLE (Night City Dark / Light Telemetry) --- */
function initTheme() {
  const toggleBtn = document.getElementById('theme-toggle');
  const storedTheme = localStorage.getItem('alief_cp_theme') || 'dark';
  setTheme(storedTheme);

  if (toggleBtn) {
    toggleBtn.addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme') || 'dark';
      const next = current === 'dark' ? 'light' : 'dark';
      setTheme(next);
    });
  }
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('alief_cp_theme', theme);

  const iconSun = document.getElementById('theme-icon-sun');
  const iconMoon = document.getElementById('theme-icon-moon');
  if (iconSun && iconMoon) {
    if (theme === 'dark') {
      iconSun.style.display = 'block';
      iconMoon.style.display = 'none';
    } else {
      iconSun.style.display = 'none';
      iconMoon.style.display = 'block';
    }
  }
}

/* --- 2. HEADER SCROLL STATE & ACTIVE LINK --- */
function initHeader() {
  const header = document.querySelector('.header');
  const sections = document.querySelectorAll('section[id]');
  const navLinks = document.querySelectorAll('.nav-link');

  const onScroll = () => {
    if (window.scrollY > 20) {
      header?.classList.add('scrolled');
    } else {
      header?.classList.remove('scrolled');
    }

    let currentId = '';
    const scrollPos = window.scrollY + 120;

    sections.forEach((sec) => {
      const top = sec.offsetTop;
      const height = sec.offsetHeight;
      if (scrollPos >= top && scrollPos < top + height) {
        currentId = sec.getAttribute('id');
      }
    });

    navLinks.forEach((link) => {
      if (link.getAttribute('href') === `#${currentId}`) {
        link.classList.add('active');
      } else {
        link.classList.remove('active');
      }
    });
  };

  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();
}

/* --- 3. READING PROGRESS INDICATOR --- */
function initReadingProgress() {
  const bar = document.getElementById('reading-progress');
  if (!bar) return;

  const update = () => {
    const scrollTop = window.scrollY || document.documentElement.scrollTop;
    const docHeight = document.documentElement.scrollHeight - window.innerHeight;
    const pct = docHeight > 0 ? (scrollTop / docHeight) * 100 : 0;
    bar.style.width = `${Math.min(100, Math.max(0, pct))}%`;
  };

  window.addEventListener('scroll', update, { passive: true });
  update();
}

/* --- 4. LIVE SURABAYA NODE TELEMETRY TIMESTAMP --- */
function initSurabayaClock() {
  const timeEl = document.getElementById('surabaya-time');
  if (!timeEl) return;

  const update = () => {
    try {
      const now = new Date();
      const formatter = new Intl.DateTimeFormat('en-US', {
        timeZone: 'Asia/Jakarta',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
      });
      timeEl.textContent = `${formatter.format(now)} WIB`;
    } catch {
      timeEl.textContent = new Date().toLocaleTimeString();
    }
  };

  update();
  setInterval(update, 1000);
}

/* --- 5. INTERACTIVE DATA-SHARDS WITH DECRYPTION TEXT SCRAMBLE --- */
const CHARS = '01#%&*+-/<>_~X$!?[];{}';

function scrambleDecrypt(element, finalText, duration = 380) {
  let start = null;
  const originalLength = finalText.length;

  function step(timestamp) {
    if (!start) start = timestamp;
    const progress = Math.min((timestamp - start) / duration, 1);
    const charsDecrypted = Math.floor(progress * originalLength);

    let output = '';
    for (let i = 0; i < originalLength; i++) {
      if (i < charsDecrypted) {
        output += finalText[i];
      } else {
        output += CHARS[Math.floor(Math.random() * CHARS.length)];
      }
    }

    element.textContent = output;

    if (progress < 1) {
      requestAnimationFrame(step);
    } else {
      element.textContent = finalText;
    }
  }

  requestAnimationFrame(step);
}

function initDataShards() {
  const shardContainers = document.querySelectorAll('.project-shards-wrapper');

  shardContainers.forEach((wrapper) => {
    const buttons = wrapper.querySelectorAll('.shard-slot-btn');
    const titleEl = wrapper.querySelector('.shard-terminal-title');
    const textEl = wrapper.querySelector('.shard-terminal-text');

    if (!buttons.length || !textEl) return;

    buttons.forEach((btn) => {
      btn.addEventListener('click', () => {
        if (btn.classList.contains('active')) return;

        buttons.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');

        const roleTitle = btn.getAttribute('data-shard-title') || 'SPEC_READOUT';
        const roleData = btn.getAttribute('data-shard-content') || '';

        if (titleEl) titleEl.textContent = `[DECRYPTED: ${roleTitle}]`;
        scrambleDecrypt(textEl, roleData, 400);
      });
    });
  });
}

/* --- 6. ROLE LENS FILTER --- */
function initRoleLens() {
  const buttons = document.querySelectorAll('[data-lens-btn]');
  const listContainer = document.getElementById('projects-list-container');
  const statusHint = document.getElementById('role-lens-status');

  if (!buttons.length || !listContainer) return;

  const descriptions = {
    all: 'Viewing full tactical dossiers across all operational roles',
    pm: 'Focusing on IT Project Management: sprint cycles, WBS, team coordination & UAT',
    ba: 'Focusing on IT Business Analysis: requirement elicitation, BPMN workflows & user stories',
    sa: 'Focusing on Systems Analysis: UML modeling, relational ERD schemas & API contracts',
    si: 'Focusing on System Implementation: ERP configuration, data cleansing, migration & SOPs'
  };

  buttons.forEach((btn) => {
    btn.addEventListener('click', () => {
      const role = btn.getAttribute('data-lens-btn');

      buttons.forEach((b) => {
        const isActive = b === btn;
        b.classList.toggle('active', isActive);
        b.setAttribute('aria-selected', isActive.toString());
      });

      listContainer.setAttribute('data-lens', role);

      // Auto-switch data-shards in projects to match the selected lens
      if (role !== 'all') {
        document.querySelectorAll(`.shard-slot-btn[data-shard-role="${role}"]`).forEach((shardBtn) => {
          shardBtn.click();
        });
      }

      // Update specialized cards
      document.querySelectorAll('.specialized-card[data-role]').forEach((card) => {
        const cardRole = card.getAttribute('data-role');
        if (role === 'all' || cardRole === role) {
          card.style.opacity = '1';
          card.style.borderColor = role === 'all' ? '' : 'var(--cp-yellow)';
        } else {
          card.style.opacity = '0.35';
          card.style.borderColor = '';
        }
      });

      if (statusHint && descriptions[role]) {
        statusHint.textContent = descriptions[role];
      }
    });
  });
}

/* --- 7. TACTICAL TERMINAL COMMAND DRAWER (QUICKHACK CONSOLE) --- */
function initTerminalDrawer() {
  const modal = document.getElementById('terminal-modal');
  const toggleBtn = document.getElementById('terminal-toggle-btn');
  const closeBtn = document.querySelector('.terminal-close-btn');
  const input = document.getElementById('terminal-input');
  const output = document.getElementById('terminal-output');

  if (!modal) return;

  const open = () => {
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';
    input?.focus();
  };

  const close = () => {
    modal.classList.remove('open');
    document.body.style.overflow = '';
  };

  toggleBtn?.addEventListener('click', open);
  closeBtn?.addEventListener('click', close);

  modal.addEventListener('click', (e) => {
    if (e.target === modal) close();
  });

  // Hotkey listener: pressing '~' or '`' toggles terminal
  window.addEventListener('keydown', (e) => {
    if (e.key === '`' || e.key === '~') {
      e.preventDefault();
      modal.classList.contains('open') ? close() : open();
    } else if (e.key === 'Escape' && modal.classList.contains('open')) {
      close();
    }
  });

  // Terminal Command Parser
  if (input && output) {
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        const cmd = input.value.trim().toLowerCase();
        input.value = '';

        const line = document.createElement('div');
        line.className = 'terminal-log-line yellow';
        line.textContent = `guest@nightcity:~$ ${cmd}`;
        output.appendChild(line);

        const resp = document.createElement('div');
        resp.className = 'terminal-log-line';

        switch (cmd) {
          case 'help':
            resp.innerHTML = `Available commands:<br>
              &bull; <strong>jump [section]</strong> — jump to hero, about, projects, experience, skills, certs, contact<br>
              &bull; <strong>cv</strong> — download latest October 2026 CV (PDF)<br>
              &bull; <strong>ppt / deck</strong> — download 12-slide executive portfolio presentation deck (PPTX)<br>
              &bull; <strong>docs</strong> — view links to BRD, FSD, and Excel QA workbooks on Google Drive<br>
              &bull; <strong>lens [pm|ba|sa|si|all]</strong> — switch active role lens<br>
              &bull; <strong>status</strong> — view system diagnostics & Surabaya node info<br>
              &bull; <strong>clear</strong> — clear console buffer<br>
              &bull; <strong>exit</strong> — close terminal`;
            break;
          case 'docs':
          case 'artifacts':
            resp.className = 'terminal-log-line cyan';
            resp.innerHTML = `VERIFIED SPECIFICATION &amp; ACADEMIC ARTIFACT ARCHIVES:<br>
              &bull; <strong>[O2C AUTOMATION CASE STUDY]</strong> Live KPI Dashboard &amp; Repository:<br>
                <a href="https://yaboialip.github.io/o2c-automation-case-study/" target="_blank" rel="noopener noreferrer" style="color:var(--cp-yellow);text-decoration:underline;">yaboialip.github.io/o2c-automation-case-study</a> // <a href="https://github.com/Yaboialip/o2c-automation-case-study" target="_blank" rel="noopener noreferrer" style="color:var(--cp-cyan);text-decoration:underline;">GitHub Repo</a><br><br>
              &bull; <strong>[NETAFARM OKR SYSTEM]</strong> BRD &amp; FSD Archive:<br>
                <a href="https://drive.google.com/drive/folders/1Ym-B6uoR0V2FTENJeCdkSvyU6m_mF0q4?usp=sharing" target="_blank" rel="noopener noreferrer" style="color:var(--cp-yellow);text-decoration:underline;">drive.google.com/.../Netafarm-BRD-FSD</a><br><br>
              &bull; <strong>[E-TEMP TELEMETRY SYSTEM]</strong> BRD, FSD &amp; Excel QA Workbook:<br>
                <a href="https://drive.google.com/drive/folders/1yJeAjFuiJ__W35Yup2n8pDP-jdXFUZht?usp=sharing" target="_blank" rel="noopener noreferrer" style="color:var(--cp-yellow);text-decoration:underline;">drive.google.com/.../ETemp-BRD-FSD-QA</a><br><br>
              &bull; <strong>[BIG DATA LAKEHOUSE]</strong> Presentation Deck (PPT):<br>
                <a href="https://drive.google.com/file/d/1-_KjfIORiWMeEm45z8R80D6LOb4IdRg4/view?usp=drive_link" target="_blank" rel="noopener noreferrer" style="color:var(--cp-yellow);text-decoration:underline;">drive.google.com/.../BigData-Deck</a><br><br>
              &bull; <strong>[UNDERGRADUATE THESIS]</strong> Full Skripsi PGEO ARIMA-GARCH:<br>
                <a href="https://drive.google.com/file/d/1sFt-QrZIYCprjvEfZtW41hoVChWjGJXn/view?usp=drive_link" target="_blank" rel="noopener noreferrer" style="color:var(--cp-yellow);text-decoration:underline;">drive.google.com/.../Skripsi-Thesis</a><br><br>
              &bull; <strong>[JOURNAL PUBLICATION]</strong> ResearchGate PDF (Smart Village IT Blueprint):<br>
                <a href="https://www.researchgate.net/publication/392726934_PERANCANGAN_IT_BLUEPRINT_SEBAGAI_STRATEGI_PENGEMBANGAN_KONSEP_SMART_VILLAGE" target="_blank" rel="noopener noreferrer" style="color:var(--cp-yellow);text-decoration:underline;">researchgate.net/.../392726934_IT_BLUEPRINT</a>`;
            break;
          case 'cv':
            resp.className = 'terminal-log-line cyan';
            resp.textContent = 'INITIATING DOWNLOAD: assets/Alief_Sukma_Dewanta_CV.pdf...';
            window.location.href = 'assets/Alief_Sukma_Dewanta_CV.pdf';
            break;
          case 'ppt':
          case 'pptx':
          case 'deck':
            resp.className = 'terminal-log-line cyan';
            resp.textContent = 'INITIATING DOWNLOAD: assets/Alief_Sukma_Dewanta_Portfolio.pptx...';
            window.location.href = 'assets/Alief_Sukma_Dewanta_Portfolio.pptx';
            break;
          case 'status':
            resp.className = 'terminal-log-line cyan';
            resp.innerHTML = `OPERATOR: ALIEF SUKMA DEWANTA<br>
              STATUS: GRADUATED OCT 2026 (TELKOM UNIVERSITY)<br>
              CONVOCATION: OCT 26, 2026 // GPA 3.66/4.00<br>
              LOCATION: SURABAYA, ID (OPEN TO RELOCATION)<br>
              CERTS: GOOGLE PM, IBM SA, SAP SCM500, SAP HR050, GCP GEN AI`;
            break;
          case 'clear':
            output.innerHTML = '';
            return;
          case 'exit':
            close();
            return;
          default:
            if (cmd.startsWith('jump ')) {
              const target = cmd.replace('jump ', '').trim();
              const el = document.getElementById(target);
              if (el) {
                close();
                el.scrollIntoView({ behavior: 'smooth' });
                resp.textContent = `NAVIGATING TO #${target}...`;
              } else {
                resp.className = 'terminal-log-line';
                resp.style.color = 'var(--cp-red)';
                resp.textContent = `TARGET NOT FOUND: #${target}. Try: hero, projects, experience, skills, contact`;
              }
            } else if (cmd.startsWith('lens ')) {
              const lens = cmd.replace('lens ', '').trim();
              const btn = document.querySelector(`[data-lens-btn="${lens}"]`);
              if (btn) {
                btn.click();
                resp.className = 'terminal-log-line cyan';
                resp.textContent = `LENS ENGAGED: ${lens.toUpperCase()}`;
              } else {
                resp.textContent = `INVALID LENS. Choose: pm, ba, sa, si, all`;
              }
            } else {
              resp.textContent = `UNKNOWN COMMAND: '${cmd}'. Type 'help' for instructions.`;
            }
        }
        output.appendChild(resp);
        output.scrollTop = output.scrollHeight;
      }
    });
  }
}

/* --- 8. LIGHTBOX MODAL --- */
function initLightbox() {
  const modal = document.getElementById('lightbox-modal');
  const modalImg = document.getElementById('lightbox-img');
  const modalTitle = document.getElementById('lightbox-title');
  const closeBtn = document.querySelector('.lightbox-close-btn');

  if (!modal || !modalImg) return;

  const open = (src, title) => {
    modalImg.src = src;
    modalImg.alt = title;
    if (modalTitle) modalTitle.textContent = title;
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';
  };

  const close = () => {
    modal.classList.remove('open');
    document.body.style.overflow = '';
  };

  document.querySelectorAll('[data-lightbox]').forEach((item) => {
    item.addEventListener('click', () => {
      const src = item.getAttribute('data-lightbox') || item.getAttribute('src');
      const title = item.getAttribute('data-title') || item.getAttribute('alt') || 'Screenshot';
      if (src) open(src, title);
    });
  });

  closeBtn?.addEventListener('click', close);

  modal.addEventListener('click', (e) => {
    if (e.target === modal) close();
  });

  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modal.classList.contains('open')) {
      close();
    }
  });
}

/* --- 9. MAGNETIC 3D PREVIEW TILT --- */
function initMediaTilt() {
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;

  const mediaContainers = document.querySelectorAll('.project-media');
  mediaContainers.forEach((media) => {
    const img = media.querySelector('img');
    if (!img) return;

    media.addEventListener('mousemove', (e) => {
      const rect = media.getBoundingClientRect();
      const x = (e.clientX - rect.left) / rect.width - 0.5;
      const y = (e.clientY - rect.top) / rect.height - 0.5;

      const rotateY = (x * 8).toFixed(2);
      const rotateX = (-y * 8).toFixed(2);

      img.style.transform = `perspective(900px) rotateX(${rotateX}deg) rotateY(${rotateY}deg) scale(1.02)`;
    });

    media.addEventListener('mouseleave', () => {
      img.style.transform = '';
    });
  });
}

/* --- 10. CONTACT FORM --- */
function initContactForm() {
  const form = document.getElementById('contact-form');
  const feedback = document.getElementById('form-feedback');

  if (!form) return;

  form.addEventListener('submit', (e) => {
    e.preventDefault();

    const name = document.getElementById('sender-name')?.value.trim();
    const email = document.getElementById('sender-email')?.value.trim();
    const message = document.getElementById('sender-message')?.value.trim();

    if (!name || !email || !message) {
      if (feedback) {
        feedback.className = 'form-feedback error';
        feedback.textContent = '[ERR] ALL FIELDS REQUIRED FOR SECURE TRANSMISSION.';
      }
      return;
    }

    const subject = encodeURIComponent(`[INQUIRY] Comm from ${name}`);
    const body = encodeURIComponent(`Hi Alief,\n\n${message}\n\nFrom: ${name} (${email})`);
    const mailtoUrl = `mailto:aliefsd84@gmail.com?subject=${subject}&body=${body}`;

    if (feedback) {
      feedback.className = 'form-feedback success';
      feedback.textContent = '[SECURE] OPENING ENCRYPTED EMAIL CLIENT...';
    }

    setTimeout(() => {
      window.location.href = mailtoUrl;
    }, 400);
  });
}

/* --- 11. MOBILE MENU --- */
function initMobileMenu() {
  const toggleBtn = document.querySelector('.mobile-toggle');
  const mobileMenu = document.getElementById('mobile-menu');
  const menuLinks = mobileMenu?.querySelectorAll('a');

  if (!toggleBtn || !mobileMenu) return;

  const toggle = () => {
    const isOpen = mobileMenu.classList.toggle('open');
    toggleBtn.setAttribute('aria-expanded', isOpen.toString());
    document.body.style.overflow = isOpen ? 'hidden' : '';
  };

  toggleBtn.addEventListener('click', toggle);

  menuLinks?.forEach((link) => {
    link.addEventListener('click', () => {
      mobileMenu.classList.remove('open');
      toggleBtn.setAttribute('aria-expanded', 'false');
      document.body.style.overflow = '';
    });
  });
}

/* --- 12. TACTICAL NAV DROPDOWN (CV & PPTX) --- */
function initNavDropdown() {
  const dropdown = document.querySelector('.nav-dropdown');
  if (!dropdown) return;

  const trigger = dropdown.querySelector('.nav-dropdown-trigger');
  if (!trigger) return;

  trigger.addEventListener('click', (e) => {
    e.stopPropagation();
    const isOpen = dropdown.classList.toggle('open');
    trigger.setAttribute('aria-expanded', isOpen.toString());
  });

  // Close when clicking outside
  document.addEventListener('click', (e) => {
    if (!dropdown.contains(e.target)) {
      dropdown.classList.remove('open');
      trigger.setAttribute('aria-expanded', 'false');
    }
  });

  // Close on Escape
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && dropdown.classList.contains('open')) {
      dropdown.classList.remove('open');
      trigger.setAttribute('aria-expanded', 'false');
    }
  });
}

