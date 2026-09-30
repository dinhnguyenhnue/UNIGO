/**
 * UNIGO — Năng lực & Phẩm chất Review Website
 * JavaScript Application Logic
 * Đảm bảo 100% logic chuẩn xác theo CV 3456, QĐ 3439, CV 5512, CT GDPT 2018
 */

// ===== Mapping Bài học -> NLS chuẩn =====
const BAC_MAP = {
    'L1-3': { bac: 1, name: 'Cơ bản 1', cb: 'CB1', lop: 'Tiền TH / Lớp 1-3' },
    'L4-5': { bac: 2, name: 'Cơ bản 2', cb: 'CB2', lop: 'Lớp 4-5' },
    'L6-7': { bac: 3, name: 'Trung cấp 1', cb: 'CB3', lop: 'Lớp 6-7' },
    'L8-9': { bac: 4, name: 'Trung cấp 2', cb: 'CB4', lop: 'Lớp 8-9' },
    'L10-12': { bac: 5, name: 'Nâng cao 1', cb: 'CB5', lop: 'Lớp 10-12' }
};

const NLS_MAPPING = {
    computer: {
        primary: '5.1',
        secondary: '4.1',
        mienPrimary: 'V',
        mienNameP: 'Giải quyết vấn đề',
        primaryName: 'Giải quyết các vấn đề kỹ thuật',
        mienSecondary: 'IV',
        mienNameS: 'An toàn',
        secondaryName: 'Bảo vệ thiết bị'
    },
    programming: {
        primary: '3.4',
        secondary: '5.3',
        mienPrimary: 'III',
        mienNameP: 'Sáng tạo nội dung số',
        primaryName: 'Lập trình',
        mienSecondary: 'V',
        mienNameS: 'Giải quyết vấn đề',
        secondaryName: 'Sử dụng sáng tạo công nghệ số'
    },
    internet: {
        primary: '1.1',
        secondary: '2.1',
        mienPrimary: 'I',
        mienNameP: 'Khai thác dữ liệu và thông tin',
        primaryName: 'Duyệt, tìm kiếm và lọc dữ liệu, thông tin và nội dung số',
        mienSecondary: 'II',
        mienNameS: 'Giao tiếp và hợp tác trong môi trường số',
        secondaryName: 'Tương tác thông qua công nghệ số'
    },
    safety: {
        primary: '4.2',
        secondary: '2.5',
        mienPrimary: 'IV',
        mienNameP: 'An toàn',
        primaryName: 'Bảo vệ dữ liệu cá nhân và quyền riêng tư',
        mienSecondary: 'II',
        mienNameS: 'Giao tiếp và hợp tác trong môi trường số',
        secondaryName: 'Quy tắc ứng xử trên mạng'
    },
    office: {
        primary: '3.1',
        secondary: '1.3',
        mienPrimary: 'III',
        mienNameP: 'Sáng tạo nội dung số',
        primaryName: 'Phát triển nội dung số',
        mienSecondary: 'I',
        mienNameS: 'Khai thác dữ liệu và thông tin',
        secondaryName: 'Quản lý dữ liệu, thông tin và nội dung số'
    },
    data: {
        primary: '1.3',
        secondary: '3.1',
        mienPrimary: 'I',
        mienNameP: 'Khai thác dữ liệu và thông tin',
        primaryName: 'Quản lý dữ liệu, thông tin và nội dung số',
        mienSecondary: 'III',
        mienNameS: 'Sáng tạo nội dung số',
        secondaryName: 'Phát triển nội dung số'
    },
    robotics: {
        primary: '5.2',
        secondary: '3.4',
        mienPrimary: 'V',
        mienNameP: 'Giải quyết vấn đề',
        primaryName: 'Xác định nhu cầu và giải pháp công nghệ',
        mienSecondary: 'III',
        mienNameS: 'Sáng tạo nội dung số',
        secondaryName: 'Lập trình'
    },
    ai: {
        primary: '6.1',
        secondary: '6.2',
        mienPrimary: 'VI',
        mienNameP: 'Ứng dụng trí tuệ nhân tạo (AI)',
        primaryName: 'Hiểu biết về trí tuệ nhân tạo (AI)',
        mienSecondary: 'VI',
        mienNameS: 'Ứng dụng trí tuệ nhân tạo (AI)',
        secondaryName: 'Sử dụng trí tuệ nhân tạo (AI)'
    }
};

const TOPIC_NAMES = {
    computer: 'Làm quen máy tính, thiết bị số',
    programming: 'Lập trình, thuật toán',
    internet: 'Internet, mạng máy tính, tìm kiếm',
    safety: 'An toàn, đạo đức, văn hóa mạng',
    office: 'Soạn thảo văn bản, trình chiếu',
    data: 'Bảng tính điện tử, quản lý dữ liệu',
    robotics: 'Robotics (Cơ chế, cảm biến, lắp ráp)',
    ai: 'Trí tuệ nhân tạo (AI), công nghệ tương lai'
};

// ===== Particles Background =====
function initParticles() {
    const container = document.getElementById('particles');
    if (!container) return;
    const particleCount = 28;
    for (let i = 0; i < particleCount; i++) {
        const particle = document.createElement('div');
        particle.className = 'particle';
        particle.style.left = Math.random() * 100 + '%';
        particle.style.top = Math.random() * 100 + '%';
        const size = Math.random() * 5 + 3;
        particle.style.width = size + 'px';
        particle.style.height = size + 'px';
        particle.style.animationDuration = (Math.random() * 14 + 10) + 's';
        particle.style.animationDelay = (Math.random() * 6) + 's';
        container.appendChild(particle);
    }
}

// ===== Navbar Scroll & Mobile Menu =====
function initNavbar() {
    const navbar = document.getElementById('navbar');
    const navToggle = document.getElementById('navToggle');
    const navMenu = document.getElementById('navMenu');
    const navLinks = document.querySelectorAll('.nav-link');

    window.addEventListener('scroll', () => {
        navbar.classList.toggle('scrolled', window.scrollY > 40);
        updateActiveNavLink();
    });

    if (navToggle) {
        navToggle.addEventListener('click', () => {
            navMenu.classList.toggle('active');
            navToggle.classList.toggle('active');
        });
    }

    navLinks.forEach(link => {
        link.addEventListener('click', () => {
            if (navMenu) navMenu.classList.remove('active');
            if (navToggle) navToggle.classList.remove('active');
        });
    });
}

function updateActiveNavLink() {
    const sections = document.querySelectorAll('section[id]');
    const scrollY = window.pageYOffset + 140;

    sections.forEach(current => {
        const sectionHeight = current.offsetHeight;
        const sectionTop = current.offsetTop;
        const sectionId = current.getAttribute('id');
        const link = document.querySelector(`.nav-link[data-section="${sectionId}"]`);

        if (scrollY > sectionTop && scrollY <= sectionTop + sectionHeight) {
            document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
            if (link) link.classList.add('active');
        }
    });
}

// ===== Count Up Animation =====
function initCountUp() {
    const counters = document.querySelectorAll('.stat-number');
    let animated = false;

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting && !animated) {
                animated = true;
                counters.forEach(counter => {
                    const target = parseInt(counter.getAttribute('data-count'), 10);
                    if (isNaN(target)) return;
                    let current = 0;
                    const step = Math.max(1, Math.floor(target / 25));
                    const timer = setInterval(() => {
                        current += step;
                        if (current >= target) {
                            counter.textContent = target;
                            clearInterval(timer);
                        } else {
                            counter.textContent = current;
                        }
                    }, 40);
                });
            }
        });
    }, { threshold: 0.3 });

    const statsElem = document.querySelector('.hero-stats');
    if (statsElem) observer.observe(statsElem);
}

// ===== Interactive Miền & Thành tố Detail Modal =====
function initMienDetailViewer() {
    // Populate direct component select dropdown
    const selectElem = document.getElementById('directComponentSelect');
    if (selectElem && typeof NLS_FULL_DATA !== 'undefined') {
        selectElem.innerHTML = '<option value="">-- Chọn 1 trong 24 Thành tố để tra cứu chi tiết --</option>';
        Object.keys(NLS_FULL_DATA).forEach(code => {
            const item = NLS_FULL_DATA[code];
            const opt = document.createElement('option');
            opt.value = code;
            opt.textContent = `${code}. ${item.name} (Miền ${NLS_MIEN_DEFINITIONS[item.mienId].roman})`;
            selectElem.appendChild(opt);
        });
    }

    // Attach click events on Miền cards
    const mienCards = document.querySelectorAll('.mien-card');
    mienCards.forEach(card => {
        const mienId = card.getAttribute('data-mien');
        card.addEventListener('click', (e) => {
            // If clicked on specific thanh to
            const ttItem = e.target.closest('.thanh-to-item');
            if (ttItem) {
                const code = ttItem.getAttribute('data-code');
                showThanhToModal(code);
            } else {
                showMienModal(mienId);
            }
        });
    });
}

function showMienModal(mienId) {
    const mien = NLS_MIEN_DEFINITIONS[mienId];
    if (!mien) return;

    let modal = document.getElementById('detailModal');
    if (!modal) {
        modal = createModalElement();
    }

    const modalBody = modal.querySelector('.modal-body');
    let thanhToListHtml = '';
    mien.thanhToCodes.forEach(code => {
        const tt = NLS_FULL_DATA[code];
        thanhToListHtml += `
            <div class="modal-tt-card" onclick="showThanhToModal('${code}')">
                <div class="modal-tt-header">
                    <span class="badge badge-blue">${code}</span>
                    <h5>${tt.name}</h5>
                </div>
                <p class="modal-tt-desc">${tt.mota}</p>
                <div class="modal-tt-action">Xem mô tả chuẩn 5 Bậc →</div>
            </div>
        `;
    });

    modalBody.innerHTML = `
        <div class="modal-header-banner">
            <span class="badge badge-blue">Miền ${mien.roman}</span>
            <h3>${mien.name}</h3>
        </div>
        <div class="modal-concept-box">
            <h4>📖 Khái niệm chính thức (CV 3456):</h4>
            <p>${mien.concept}</p>
        </div>
        <h4 style="margin: 20px 0 12px; color: var(--text-primary); font-weight: 700;">Các Thành tố trực thuộc (${mien.thanhToCodes.length} thành tố):</h4>
        <div class="modal-tt-grid">
            ${thanhToListHtml}
        </div>
    `;

    modal.classList.add('active');
}

function showThanhToModal(code) {
    const tt = NLS_FULL_DATA[code];
    if (!tt) return;
    const mien = NLS_MIEN_DEFINITIONS[tt.mienId];

    let modal = document.getElementById('detailModal');
    if (!modal) {
        modal = createModalElement();
    }

    const modalBody = modal.querySelector('.modal-body');
    modalBody.innerHTML = `
        <div class="modal-header-banner">
            <span class="badge badge-blue">Thành tố ${code}</span>
            <span class="badge badge-purple">Miền ${mien.roman}. ${mien.name}</span>
            <h3>${tt.name}</h3>
        </div>
        <div class="modal-concept-box">
            <h4>📖 Khái niệm / Mô tả chuẩn từ CV 3456:</h4>
            <p>${tt.mota}</p>
        </div>
        <div style="margin-top: 20px;">
            <h4 style="margin-bottom: 12px; color: var(--text-primary); font-weight: 700;">🎯 Bảng Descriptors (Yêu cầu cần đạt) theo 5 Bậc:</h4>
            <div class="table-responsive">
                <table class="styled-table" style="font-size: 0.92rem;">
                    <thead>
                        <tr>
                            <th style="width: 140px;">Khối lớp & Bậc</th>
                            <th>Mô tả chuẩn Yêu cầu cần đạt (CV 3456)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><strong>Lớp 1-3</strong><br><span class="badge badge-blue">Bậc 1 (CB1)</span></td>
                            <td style="white-space: pre-line;">${tt.descriptors['L1-3']}</td>
                        </tr>
                        <tr>
                            <td><strong>Lớp 4-5</strong><br><span class="badge badge-teal">Bậc 2 (CB2)</span></td>
                            <td style="white-space: pre-line;">${tt.descriptors['L4-5']}</td>
                        </tr>
                        <tr class="highlight-row">
                            <td><strong>Lớp 6-7</strong><br><span class="badge badge-gold">Bậc 3 (CB3)</span></td>
                            <td style="white-space: pre-line;">${tt.descriptors['L6-7']}</td>
                        </tr>
                        <tr class="highlight-row">
                            <td><strong>Lớp 8-9</strong><br><span class="badge badge-orange">Bậc 4 (CB4)</span></td>
                            <td style="white-space: pre-line;">${tt.descriptors['L8-9']}</td>
                        </tr>
                        <tr>
                            <td><strong>Lớp 10-12</strong><br><span class="badge badge-purple">Bậc 5 (CB5)</span></td>
                            <td style="white-space: pre-line;">${tt.descriptors['L10-12']}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    `;

    modal.classList.add('active');
}

function createModalElement() {
    const modal = document.createElement('div');
    modal.id = 'detailModal';
    modal.className = 'custom-modal';
    modal.innerHTML = `
        <div class="modal-overlay" onclick="closeModal()"></div>
        <div class="modal-dialog">
            <button class="modal-close" onclick="closeModal()">✕</button>
            <div class="modal-body"></div>
        </div>
    `;
    document.body.appendChild(modal);
    return modal;
}

function closeModal() {
    const modal = document.getElementById('detailModal');
    if (modal) modal.classList.remove('active');
}

// Direct Component Lookup Select Handler
function onDirectComponentChange() {
    const selectElem = document.getElementById('directComponentSelect');
    const resultDiv = document.getElementById('directComponentResult');
    if (!selectElem || !resultDiv) return;
    const code = selectElem.value;
    if (!code) {
        resultDiv.innerHTML = '';
        return;
    }
    const tt = NLS_FULL_DATA[code];
    if (!tt) return;
    const mien = NLS_MIEN_DEFINITIONS[tt.mienId];

    resultDiv.innerHTML = `
        <div class="result-card mt-3">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
                <h4>📌 Thành tố ${tt.code}: ${tt.name}</h4>
                <span class="badge badge-blue">Miền ${mien.roman}. ${mien.name}</span>
            </div>
            <div class="result-item" style="background: rgba(37,99,235,0.05); padding: 12px; border-radius: 8px; margin: 12px 0;">
                <span class="result-label" style="font-weight: 700; color: var(--accent-blue);">📖 Khái niệm chính thức theo CV 3456:</span><br>
                <div style="color: var(--text-primary); font-size: 0.98rem; line-height: 1.6; margin-top: 6px;">
                    ${tt.mota}
                </div>
            </div>
            <div style="margin-top: 14px;">
                <h5 style="margin-bottom: 8px; font-weight: 700; color: var(--text-secondary);">Mô tả chuẩn YCCD theo từng Bậc:</h5>
                <div class="descriptor-tabs">
                    <button class="desc-tab-btn active" onclick="switchDescTab(this, 'L1-3')">Bậc 1 (L1-3)</button>
                    <button class="desc-tab-btn" onclick="switchDescTab(this, 'L4-5')">Bậc 2 (L4-5)</button>
                    <button class="desc-tab-btn" onclick="switchDescTab(this, 'L6-7')">Bậc 3 (L6-7)</button>
                    <button class="desc-tab-btn" onclick="switchDescTab(this, 'L8-9')">Bậc 4 (L8-9)</button>
                    <button class="desc-tab-btn" onclick="switchDescTab(this, 'L10-12')">Bậc 5 (L10-12)</button>
                </div>
                <div id="descTabContent" class="desc-tab-box">
                    <pre style="white-space: pre-line; font-family: var(--font-main); font-size: 0.95rem; color: var(--text-primary);">${tt.descriptors['L1-3']}</pre>
                </div>
            </div>
        </div>
    `;
}

function switchDescTab(btn, gradeKey) {
    document.querySelectorAll('.desc-tab-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const selectElem = document.getElementById('directComponentSelect');
    const code = selectElem ? selectElem.value : null;
    if (code && NLS_FULL_DATA[code]) {
        const contentBox = document.getElementById('descTabContent');
        if (contentBox) {
            contentBox.innerHTML = `
                <pre style="white-space: pre-line; font-family: var(--font-main); font-size: 0.95rem; color: var(--text-primary);">${NLS_FULL_DATA[code].descriptors[gradeKey]}</pre>
            `;
        }
    }
}

// ===== NLS Lookup Tool by Class & Topic =====
function lookupNLS() {
    const classKey = document.getElementById('lookupClass').value;
    const topicKey = document.getElementById('lookupTopic').value;
    const resultDiv = document.getElementById('lookupResult');
    if (!resultDiv) return;

    const bac = BAC_MAP[classKey];
    const mapping = NLS_MAPPING[topicKey];
    const topicName = TOPIC_NAMES[topicKey];

    const primaryData = NLS_FULL_DATA[mapping.primary];
    const secondaryData = NLS_FULL_DATA[mapping.secondary];

    const primaryDesc = primaryData ? primaryData.descriptors[classKey] : '';
    const secondaryDesc = secondaryData ? secondaryData.descriptors[classKey] : '';

    resultDiv.innerHTML = `
        <div class="result-card">
            <h4>📋 Kết quả tra cứu theo chuẩn CV 3456</h4>
            
            <div class="result-item">
                <span class="result-label">Nhóm bài học:</span> <strong>${topicName}</strong>
            </div>
            
            <div class="result-item">
                <span class="result-label">Khối lớp áp dụng:</span> ${bac.lop} → <strong>Bậc ${bac.bac}</strong> (${bac.name}) → Ký hiệu chuẩn: <code>${bac.cb}</code>
            </div>

            <div class="result-box-highlight" style="margin: 14px 0; padding: 14px; background: rgba(37,99,235,0.06); border-left: 4px solid var(--accent-blue); border-radius: 6px;">
                <div style="font-weight: 700; color: var(--accent-blue); font-size: 1rem; margin-bottom: 6px;">
                    🎯 Miền NLS chính: Miền ${mapping.mienPrimary}. ${mapping.mienNameP}
                </div>
                <div style="font-weight: 600; color: var(--text-primary); margin-bottom: 4px;">
                    Thành tố <code>${mapping.primary}</code>: ${mapping.primaryName}
                </div>
                <div style="font-size: 0.92rem; color: var(--text-secondary); margin-bottom: 8px;">
                    <strong>Khái niệm cụ thể (CV 3456):</strong> ${primaryData ? primaryData.mota : ''}
                </div>
                <div style="font-size: 0.92rem; color: var(--text-primary); background: #ffffff; padding: 10px; border-radius: 6px; border: 1px solid var(--border-subtle);">
                    <strong>Descriptor chuẩn Bậc ${bac.bac}:</strong><br>
                    <span style="white-space: pre-line; line-height: 1.6;">${primaryDesc}</span>
                </div>
            </div>

            <div class="result-box-highlight" style="margin: 14px 0; padding: 14px; background: rgba(16,185,129,0.06); border-left: 4px solid var(--accent-green); border-radius: 6px;">
                <div style="font-weight: 700; color: var(--accent-green); font-size: 1rem; margin-bottom: 6px;">
                    💡 Miền NLS phụ: Miền ${mapping.mienSecondary}. ${mapping.mienNameS}
                </div>
                <div style="font-weight: 600; color: var(--text-primary); margin-bottom: 4px;">
                    Thành tố <code>${mapping.secondary}</code>: ${mapping.secondaryName}
                </div>
                <div style="font-size: 0.92rem; color: var(--text-secondary); margin-bottom: 8px;">
                    <strong>Khái niệm cụ thể (CV 3456):</strong> ${secondaryData ? secondaryData.mota : ''}
                </div>
                <div style="font-size: 0.92rem; color: var(--text-primary); background: #ffffff; padding: 10px; border-radius: 6px; border: 1px solid var(--border-subtle);">
                    <strong>Descriptor chuẩn Bậc ${bac.bac}:</strong><br>
                    <span style="white-space: pre-line; line-height: 1.6;">${secondaryDesc}</span>
                </div>
            </div>

            <div class="result-item" style="margin-top: 14px; padding-top: 12px; border-top: 1px solid var(--border-subtle);">
                <span class="result-label">Cú pháp viết mã CB nhanh:</span>
                <code>${mapping.primary}.${bac.cb}a</code> (Chính) &nbsp;|&nbsp;
                <code>${mapping.secondary}.${bac.cb}a</code> (Phụ)
            </div>

            <div class="result-item" style="border-bottom: none; margin-top: 10px;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                    <span class="result-label" style="font-weight: 700;">Mẫu soạn hoàn chỉnh vào mục 2.2 KHBD:</span>
                    <button class="copy-btn" onclick="copyLookupTemplate(this)" style="padding: 4px 10px; font-size: 0.82rem;">📋 Copy mẫu</button>
                </div>
                <div class="code-block" style="margin: 0;">
                    <pre id="lookupCodePre"><code style="font-size: 0.9rem; line-height: 1.6;">2.2. Năng lực số (Thông tư 02/2025 – CV 3456):
- Miền ${mapping.mienPrimary}. ${mapping.mienNameP} (thành tố ${mapping.primary}. ${mapping.primaryName} – Bậc ${bac.bac}):
  ${primaryDesc.replace(/\n/g, '\n  ')}
  (Đạt được thông qua Hoạt động 2, Hoạt động 3).
- Miền ${mapping.mienSecondary}. ${mapping.mienNameS} (thành tố ${mapping.secondary}. ${mapping.secondaryName} – Bậc ${bac.bac}):
  ${secondaryDesc.replace(/\n/g, '\n  ')}
  (Đạt được thông qua Hoạt động 3, Hoạt động 4).</code></pre>
                </div>
            </div>
        </div>
    `;
}

function copyLookupTemplate(btn) {
    const pre = document.getElementById('lookupCodePre');
    if (!pre) return;
    navigator.clipboard.writeText(pre.innerText).then(() => {
        btn.textContent = '✅ Đã chép!';
        setTimeout(() => { btn.textContent = '📋 Copy mẫu'; }, 2000);
    });
}

// ===== QUIZ SYSTEM OVERHAUL =====
let currentQuizDoc = 'cv3456';
let currentQuizSub = 'cv3456_bac';
let activeQuizQuestions = [];
let quizCurrentIndex = 0;
let quizAnswers = [];
let quizFinished = false;

function initQuizSystem() {
    renderQuizTabs();
    selectQuizSub(currentQuizDoc, currentQuizSub);
}

function renderQuizTabs() {
    const docTabsContainer = document.getElementById('quizDocTabs');
    if (!docTabsContainer) return;

    docTabsContainer.innerHTML = `
        <button class="quiz-doc-btn ${currentQuizDoc === 'cv3456' ? 'active' : ''}" onclick="selectQuizDoc('cv3456')">
            💻 CV 3456 (Năng lực số)
        </button>
        <button class="quiz-doc-btn ${currentQuizDoc === 'qd3439' ? 'active' : ''}" onclick="selectQuizDoc('qd3439')">
            🤖 QĐ 3439 (Năng lực AI)
        </button>
        <button class="quiz-doc-btn ${currentQuizDoc === 'cv5512' ? 'active' : ''}" onclick="selectQuizDoc('cv5512')">
            ⭐ CV 5512 & GDPT 2018 (NL Chung & 5 Phẩm chất)
        </button>
        <button class="quiz-doc-btn ${currentQuizDoc === 'dacthu' ? 'active' : ''}" onclick="selectQuizDoc('dacthu')">
            ⚙️ NL Đặc thù (Tin học & Robotics)
        </button>
        <button class="quiz-doc-btn quiz-doc-all ${currentQuizDoc === 'all' ? 'active' : ''}" onclick="selectQuizDoc('all')">
            🏆 Thi thử Toàn diện (20 câu tổng hợp)
        </button>
    `;

    renderQuizSubTabs();
}

function renderQuizSubTabs() {
    const subTabsContainer = document.getElementById('quizSubTabs');
    if (!subTabsContainer) return;

    if (currentQuizDoc === 'all') {
        subTabsContainer.innerHTML = `
            <div class="quiz-sub-desc">
                🎯 <strong>Chế độ Thi thử Toàn diện:</strong> Hệ thống tự động chọn 20 câu hỏi ngẫu nhiên từ toàn bộ các văn bản (CV 3456, QĐ 3439, CV 5512, CT GDPT 2018 và NL Đặc thù).
            </div>
        `;
        return;
    }

    const doc = QUIZ_BANK[currentQuizDoc];
    if (!doc) return;

    let html = '<div class="quiz-sub-pills">';
    Object.keys(doc.subsections).forEach(subKey => {
        const sub = doc.subsections[subKey];
        const isActive = subKey === currentQuizSub;
        html += `
            <button class="quiz-sub-pill ${isActive ? 'active' : ''}" onclick="selectQuizSub('${currentQuizDoc}', '${subKey}')">
                ${sub.name} <span class="pill-count">(${sub.questions.length} câu)</span>
            </button>
        `;
    });
    html += '</div>';
    subTabsContainer.innerHTML = html;
}

function selectQuizDoc(docKey) {
    currentQuizDoc = docKey;
    if (docKey === 'all') {
        currentQuizSub = 'all';
        activeQuizQuestions = shuffleArray(getAllQuestions()).slice(0, 20);
    } else {
        const firstSub = Object.keys(QUIZ_BANK[docKey].subsections)[0];
        currentQuizSub = firstSub;
        activeQuizQuestions = QUIZ_BANK[docKey].subsections[firstSub].questions;
    }
    renderQuizTabs();
    startNewQuizSession();
}

function selectQuizSub(docKey, subKey) {
    currentQuizDoc = docKey;
    currentQuizSub = subKey;
    activeQuizQuestions = QUIZ_BANK[docKey].subsections[subKey].questions;
    renderQuizTabs();
    startNewQuizSession();
}

function startNewQuizSession() {
    quizCurrentIndex = 0;
    quizAnswers = new Array(activeQuizQuestions.length).fill(-1);
    quizFinished = false;

    const scoreDiv = document.getElementById('quizScore');
    if (scoreDiv) scoreDiv.style.display = 'none';

    renderQuizQuestion();
}

function renderQuizQuestion() {
    const total = activeQuizQuestions.length;
    if (total === 0) return;
    const q = activeQuizQuestions[quizCurrentIndex];

    // Progress bar
    const bar = document.getElementById('quizProgressBar');
    if (bar) bar.style.width = ((quizCurrentIndex + 1) / total * 100) + '%';

    // Topic Header Badge
    let topicTitle = '';
    if (currentQuizDoc === 'all') {
        topicTitle = `🏆 Đề tổng hợp toàn diện — ${q.docTitle || ''}`;
    } else {
        const doc = QUIZ_BANK[currentQuizDoc];
        const sub = doc.subsections[currentQuizSub];
        topicTitle = `${doc.badge} • ${sub.name}`;
    }

    const content = document.getElementById('quizContent');
    content.innerHTML = `
        <div class="quiz-question-header">
            <div class="quiz-topic-badge">${topicTitle}</div>
            <div class="quiz-counter">Câu ${quizCurrentIndex + 1} / ${total}</div>
        </div>
        <div class="quiz-question">
            <h4>${q.question}</h4>
            <div class="quiz-options">
                ${q.options.map((opt, i) => `
                    <div class="quiz-option ${quizAnswers[quizCurrentIndex] === i ? 'selected' : ''}" 
                         onclick="selectQuizAnswer(${i})" data-index="${i}">
                        <span class="opt-prefix">${String.fromCharCode(65 + i)}</span>
                        <span class="opt-text">${opt}</span>
                    </div>
                `).join('')}
            </div>
            ${quizAnswers[quizCurrentIndex] >= 0 ? `
                <div class="quiz-instant-feedback ${quizAnswers[quizCurrentIndex] === q.correct ? 'correct' : 'incorrect'}">
                    <strong>${quizAnswers[quizCurrentIndex] === q.correct ? '✅ Chính xác!' : '❌ Chưa chính xác!'}</strong>
                    <div style="margin-top: 4px; font-size: 0.93rem;">${q.explanation}</div>
                </div>
            ` : ''}
        </div>
    `;

    // Button controls
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    const submitBtn = document.getElementById('submitBtn');
    const restartBtn = document.getElementById('restartBtn');

    if (prevBtn) prevBtn.style.display = quizCurrentIndex > 0 ? 'inline-flex' : 'none';
    if (nextBtn) nextBtn.style.display = quizCurrentIndex < total - 1 ? 'inline-flex' : 'none';
    if (submitBtn) submitBtn.style.display = quizCurrentIndex === total - 1 ? 'inline-flex' : 'none';
    if (restartBtn) restartBtn.style.display = 'none';
}

function selectQuizAnswer(index) {
    if (quizFinished) return;
    quizAnswers[quizCurrentIndex] = index;
    renderQuizQuestion();
}

function nextQuestion() {
    if (quizCurrentIndex < activeQuizQuestions.length - 1) {
        quizCurrentIndex++;
        renderQuizQuestion();
    }
}

function prevQuestion() {
    if (quizCurrentIndex > 0) {
        quizCurrentIndex--;
        renderQuizQuestion();
    }
}

function submitQuiz() {
    quizFinished = true;
    let correct = 0;
    activeQuizQuestions.forEach((q, i) => {
        if (quizAnswers[i] === q.correct) correct++;
    });

    const total = activeQuizQuestions.length;
    const percentage = Math.round((correct / total) * 100);

    let message = '';
    if (percentage >= 90) message = '🏆 Xuất sắc! Thầy/Cô đã nắm vững kiến thức chuẩn của văn bản!';
    else if (percentage >= 70) message = '👍 Rất tốt! Thầy/Cô chỉ cần lưu ý lại một vài điểm nhỏ.';
    else if (percentage >= 50) message = '📖 Đạt yêu cầu cơ bản. Thầy/Cô nên xem lại các câu giải thích chi tiết bên dưới.';
    else message = '📚 Cần ôn tập thêm! Thầy/Cô vui lòng đọc kỹ lại tài liệu và làm lại bài kiểm tra.';

    const content = document.getElementById('quizContent');
    content.innerHTML = '';

    const scoreDiv = document.getElementById('quizScore');
    scoreDiv.style.display = 'block';
    scoreDiv.innerHTML = `
        <div class="score-card-header">
            <div class="score-circle">
                <span class="score-number">${correct}/${total}</span>
                <span class="score-pct">${percentage}%</span>
            </div>
            <div class="score-details">
                <h3>Kết quả kiểm tra</h3>
                <p class="score-message">${message}</p>
            </div>
        </div>

        <div class="score-review-list mt-3">
            <h4 style="margin-bottom: 16px; color: var(--text-primary); font-weight: 700;">📝 Bảng đối chiếu câu hỏi & Trích dẫn văn bản:</h4>
            ${activeQuizQuestions.map((q, i) => {
                const isCorrect = quizAnswers[i] === q.correct;
                const userOpt = quizAnswers[i] >= 0 ? `${String.fromCharCode(65 + quizAnswers[i])}. ${q.options[quizAnswers[i]]}` : 'Chưa chọn đáp án';
                const correctOpt = `${String.fromCharCode(65 + q.correct)}. ${q.options[q.correct]}`;
                return `
                    <div class="review-item ${isCorrect ? 'review-correct' : 'review-wrong'}">
                        <div class="review-title">
                            <span>${isCorrect ? '✅' : '❌'} Câu ${i + 1}: ${q.question}</span>
                        </div>
                        ${!isCorrect ? `
                            <div class="review-row wrong-text">
                                <strong>Lựa chọn của Thầy/Cô:</strong> ${userOpt}
                            </div>
                        ` : ''}
                        <div class="review-row correct-text">
                            <strong>Đáp án chuẩn văn bản:</strong> ${correctOpt}
                        </div>
                        <div class="review-explanation">
                            <strong>Trích dẫn căn cứ:</strong> ${q.explanation}
                        </div>
                    </div>
                `;
            }).join('')}
        </div>
    `;

    document.getElementById('prevBtn').style.display = 'none';
    document.getElementById('nextBtn').style.display = 'none';
    document.getElementById('submitBtn').style.display = 'none';
    document.getElementById('restartBtn').style.display = 'inline-flex';
    document.getElementById('quizProgressBar').style.width = '100%';
}

function restartQuiz() {
    startNewQuizSession();
}

function shuffleArray(array) {
    const arr = [...array];
    for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
}

// ===== Code Copy Helper =====
function copyCode(btn) {
    const codeBlock = btn.closest('.code-block').querySelector('code');
    if (!codeBlock) return;
    navigator.clipboard.writeText(codeBlock.innerText).then(() => {
        btn.textContent = '✅ Đã copy!';
        setTimeout(() => { btn.textContent = '📋 Copy'; }, 2000);
    });
}

// ===== Tab Switching Helpers =====
function initTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const parent = btn.closest('.card') || btn.closest('.tab-container') || document;
            parent.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            parent.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

            btn.classList.add('active');
            const targetId = btn.getAttribute('data-tab');
            const targetContent = document.getElementById(targetId);
            if (targetContent) targetContent.classList.add('active');
        });
    });
}

// ===== Subject Toggle (Tin học / Robotics) =====
function initSubjectToggle() {
    const toggleBtns = document.querySelectorAll('.toggle-btn');
    toggleBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            toggleBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            const subject = btn.getAttribute('data-subject');
            document.querySelectorAll('.subject-panel').forEach(p => p.classList.remove('active'));
            const activePanel = document.getElementById(`panel-${subject}`);
            if (activePanel) activePanel.classList.add('active');
        });
    });
}

// ===== Back to Top =====
function initBackToTop() {
    const btn = document.getElementById('backToTop');
    if (!btn) return;
    window.addEventListener('scroll', () => {
        btn.classList.toggle('visible', window.scrollY > 450);
    });
    btn.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
}

// ===== Global Init on DOM Ready =====
document.addEventListener('DOMContentLoaded', () => {
    initParticles();
    initNavbar();
    initCountUp();
    initTabs();
    initSubjectToggle();
    initBackToTop();
    initMienDetailViewer();
    initQuizSystem();

    // Trigger initial lookup
    if (document.getElementById('lookupClass')) {
        lookupNLS();
    }
});
