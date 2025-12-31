/**
 * cpyhwpx API Documentation App
 */

class ApiDocApp {
    constructor() {
        this.data = null;
        this.currentCategory = 'all';
        this.currentFunction = null;

        // DOM 요소 캐싱
        this.elements = {
            searchInput: document.getElementById('search-input'),
            searchResults: document.getElementById('search-results'),
            categoryList: document.getElementById('category-list'),
            functionList: document.getElementById('function-list'),
            functionCount: document.getElementById('function-count'),
            welcomeView: document.getElementById('welcome-view'),
            functionDetail: document.getElementById('function-detail'),
            relatedPanel: document.getElementById('related-panel'),
            dependencySection: document.getElementById('dependency-section'),
            dependencyList: document.getElementById('dependency-list'),
            sameCategoryList: document.getElementById('same-category-list'),
            sidebar: document.getElementById('sidebar'),
            sidebarOverlay: document.getElementById('sidebar-overlay'),
            mobileMenuBtn: document.querySelector('.mobile-menu-btn')
        };
    }

    /**
     * 앱 초기화
     */
    async init() {
        try {
            // 데이터 로드
            await this.loadData();

            // UI 초기화
            this.renderCategories();
            this.renderFunctionList();
            this.renderStats();

            // 이벤트 리스너 설정
            this.setupEventListeners();

            // URL 해시 처리
            this.handleHashChange();

        } catch (error) {
            console.error('Failed to initialize:', error);
            this.showError('데이터를 로드하는 중 오류가 발생했습니다.');
        }
    }

    /**
     * API 데이터 로드
     */
    async loadData() {
        const response = await fetch('data/api.json');
        if (!response.ok) {
            throw new Error('Failed to load API data');
        }
        this.data = await response.json();

        // 검색 엔진에 데이터 설정
        window.searchEngine.setData(this.data);
    }

    /**
     * 카테고리 렌더링
     */
    renderCategories() {
        const container = this.elements.categoryList;
        let html = `<li class="active" data-category="all">
            전체 <span class="cat-count">${this.data.total_functions}</span>
        </li>`;

        this.data.categories.forEach(cat => {
            html += `<li data-category="${cat.id}">
                ${cat.name} <span class="cat-count">${cat.count}</span>
            </li>`;
        });

        container.innerHTML = html;
    }

    /**
     * 함수 목록 렌더링
     */
    renderFunctionList(filter = null) {
        const container = this.elements.functionList;
        let functions = window.searchEngine.filterByCategory(this.currentCategory);

        // 검색 필터
        if (filter) {
            functions = window.searchEngine.search(filter, {
                category: this.currentCategory
            });
        }

        // 함수 개수 표시
        this.elements.functionCount.textContent = functions.length;

        // 목록 렌더링
        let html = '';
        functions.forEach(func => {
            const isActive = this.currentFunction && this.currentFunction.name === func.name;
            html += `<li data-function="${func.name}" class="${isActive ? 'active' : ''}">
                ${func.name}
            </li>`;
        });

        container.innerHTML = html || '<li class="no-results">함수를 찾을 수 없습니다</li>';
    }

    /**
     * 통계 렌더링
     */
    renderStats() {
        document.getElementById('stat-total').textContent = this.data.total_functions;
        document.getElementById('stat-categories').textContent = this.data.categories.length;

        // 상세 문서화된 함수 수
        const detailed = this.data.functions.filter(f => f.is_detailed).length;
        document.getElementById('stat-detailed').textContent = detailed;
    }

    /**
     * 함수 상세 정보 렌더링
     */
    renderFunctionDetail(funcName) {
        const func = this.data.functions.find(f => f.name === funcName);
        if (!func) {
            this.showWelcome();
            return;
        }

        this.currentFunction = func;
        this.elements.welcomeView.style.display = 'none';
        this.elements.functionDetail.style.display = 'block';

        // 카테고리 정보
        const category = this.data.categories.find(c => c.id === func.category) || { name: '기타' };

        // HTML 생성
        let html = `
            <div class="function-header">
                <h1>${func.name}</h1>
                <span class="category-badge">${category.name}</span>
            </div>

            <div class="function-signature">
                <code><span class="sig-func">hwp.${func.name}</span>(${this.formatSignature(func.args)})</code>
            </div>

            <div class="function-description">
                ${func.description || '<em>설명이 없습니다.</em>'}
            </div>

            ${this.renderDependencies(func)}
            ${this.renderArgs(func.args)}
            ${this.renderReturns(func.returns)}
            ${this.renderExamples(func.examples)}
        `;

        this.elements.functionDetail.innerHTML = html;

        // 연관 함수 패널 업데이트
        this.renderRelatedPanel(func);

        // 함수 목록에서 활성화
        this.updateFunctionListActive(funcName);
    }

    /**
     * 시그니처 포맷팅
     */
    formatSignature(args) {
        if (!args || args.length === 0) return '';

        return args.map(arg => {
            if (arg.default !== null && arg.default !== undefined) {
                return `<span class="sig-param">${arg.name}</span>=<span class="sig-default">${arg.default}</span>`;
            }
            return `<span class="sig-param">${arg.name}</span>`;
        }).join(', ');
    }

    /**
     * 선행 조건 렌더링
     */
    renderDependencies(func) {
        if ((!func.requires || func.requires.length === 0) &&
            (!func.before || func.before.length === 0)) {
            return '';
        }

        let html = '<div class="dependencies-section"><h2>선행 조건</h2>';

        if (func.requires && func.requires.length > 0) {
            func.requires.forEach(req => {
                html += `<div class="dependency-item">
                    <span class="dependency-label">@requires</span>
                    <a href="#${req}" class="dependency-link">${req}</a>
                </div>`;
            });
        }

        if (func.before && func.before.length > 0) {
            func.before.forEach(before => {
                html += `<div class="dependency-item">
                    <span class="dependency-label">@before</span>
                    <a href="#${before}" class="dependency-link">${before}</a>
                </div>`;
            });
        }

        html += '</div>';
        return html;
    }

    /**
     * 인자 렌더링
     */
    renderArgs(args) {
        if (!args || args.length === 0) return '';

        let html = '<div class="detail-section"><h2>매개변수 (Args)</h2>';
        html += '<table class="args-table"><thead><tr>';
        html += '<th>이름</th><th>기본값</th><th>설명</th>';
        html += '</tr></thead><tbody>';

        args.forEach(arg => {
            html += `<tr>
                <td class="arg-name">${arg.name}</td>
                <td class="arg-default">${arg.default !== null && arg.default !== undefined ? arg.default : '-'}</td>
                <td class="arg-desc">
                    ${arg.description || '-'}
                    ${arg.options && arg.options.length > 0 ? this.renderArgOptions(arg.options) : ''}
                </td>
            </tr>`;
        });

        html += '</tbody></table></div>';
        return html;
    }

    /**
     * 인자 옵션 렌더링
     */
    renderArgOptions(options) {
        if (!options || options.length === 0) return '';

        let html = '<ul class="arg-options">';
        options.forEach(opt => {
            html += `<li>${opt}</li>`;
        });
        html += '</ul>';
        return html;
    }

    /**
     * 반환값 렌더링
     */
    renderReturns(returns) {
        if (!returns) return '';

        return `<div class="detail-section">
            <h2>반환값 (Returns)</h2>
            <div class="returns-content">${returns}</div>
        </div>`;
    }

    /**
     * 예제 렌더링
     */
    renderExamples(examples) {
        if (!examples || examples.length === 0) return '';

        const code = examples.join('\n');
        return `<div class="detail-section">
            <h2>예제 (Examples)</h2>
            <div class="examples-content">
                <pre>${this.escapeHtml(code)}</pre>
            </div>
        </div>`;
    }

    /**
     * 연관 함수 패널 렌더링
     */
    renderRelatedPanel(func) {
        const related = window.searchEngine.findRelated(func, 5);

        // 의존성 섹션
        const dependencies = related.filter(r => r.relationType === 'requires' || r.relationType === 'before');
        if (dependencies.length > 0) {
            this.elements.dependencySection.style.display = 'block';
            let html = '';
            dependencies.forEach(dep => {
                const label = dep.relationType === 'requires' ? '필수' : '권장';
                html += `<a href="#${dep.name}" title="${label}">${dep.name}</a>`;
            });
            this.elements.dependencyList.innerHTML = html;
        } else {
            this.elements.dependencySection.style.display = 'none';
        }

        // 같은 카테고리
        const sameCategory = related.filter(r => r.relationType === 'category');
        let html = '';
        sameCategory.forEach(f => {
            html += `<a href="#${f.name}">${f.name}</a>`;
        });
        this.elements.sameCategoryList.innerHTML = html || '<span style="color: var(--text-muted); font-size: 13px;">없음</span>';

        // 패널 표시
        this.elements.relatedPanel.classList.add('active');
    }

    /**
     * 환영 화면 표시
     */
    showWelcome() {
        this.currentFunction = null;
        this.elements.welcomeView.style.display = 'block';
        this.elements.functionDetail.style.display = 'none';
        this.elements.relatedPanel.classList.remove('active');
        this.updateFunctionListActive(null);
    }

    /**
     * 함수 목록 활성화 상태 업데이트
     */
    updateFunctionListActive(funcName) {
        const items = this.elements.functionList.querySelectorAll('li');
        items.forEach(item => {
            if (item.dataset.function === funcName) {
                item.classList.add('active');
            } else {
                item.classList.remove('active');
            }
        });
    }

    /**
     * 이벤트 리스너 설정
     */
    setupEventListeners() {
        // 카테고리 클릭
        this.elements.categoryList.addEventListener('click', (e) => {
            const li = e.target.closest('li');
            if (!li) return;

            this.currentCategory = li.dataset.category;

            // 활성화 상태 변경
            this.elements.categoryList.querySelectorAll('li').forEach(item => {
                item.classList.remove('active');
            });
            li.classList.add('active');

            // 함수 목록 업데이트
            this.renderFunctionList(this.elements.searchInput.value);
        });

        // 함수 클릭
        this.elements.functionList.addEventListener('click', (e) => {
            const li = e.target.closest('li');
            if (!li || !li.dataset.function) return;

            window.location.hash = li.dataset.function;
        });

        // 검색 입력
        this.elements.searchInput.addEventListener('input', (e) => {
            const query = e.target.value.trim();
            this.handleSearch(query);
        });

        // 검색 포커스
        this.elements.searchInput.addEventListener('focus', () => {
            const query = this.elements.searchInput.value.trim();
            if (query.length >= 1) {
                this.showSearchResults(query);
            }
        });

        // 검색 결과 클릭
        this.elements.searchResults.addEventListener('click', (e) => {
            const item = e.target.closest('.search-result-item');
            if (item) {
                window.location.hash = item.dataset.function;
                this.elements.searchResults.classList.remove('active');
                this.elements.searchInput.value = '';
                this.renderFunctionList();
            }
        });

        // 문서 클릭 시 검색 결과 닫기
        document.addEventListener('click', (e) => {
            if (!this.elements.searchResults.contains(e.target) &&
                e.target !== this.elements.searchInput) {
                this.elements.searchResults.classList.remove('active');
            }
        });

        // 해시 변경
        window.addEventListener('hashchange', () => this.handleHashChange());

        // 모바일 메뉴
        if (this.elements.mobileMenuBtn) {
            this.elements.mobileMenuBtn.addEventListener('click', () => {
                this.elements.sidebar.classList.toggle('open');
                this.elements.sidebarOverlay.classList.toggle('active');
            });
        }

        // 사이드바 오버레이 클릭
        this.elements.sidebarOverlay.addEventListener('click', () => {
            this.elements.sidebar.classList.remove('open');
            this.elements.sidebarOverlay.classList.remove('active');
        });

        // 키보드 단축키
        document.addEventListener('keydown', (e) => {
            // Ctrl/Cmd + K: 검색 포커스
            if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
                e.preventDefault();
                this.elements.searchInput.focus();
            }
            // Escape: 검색 결과 닫기
            if (e.key === 'Escape') {
                this.elements.searchResults.classList.remove('active');
                this.elements.searchInput.blur();
            }
        });
    }

    /**
     * 검색 처리
     */
    handleSearch(query) {
        if (query.length < 1) {
            this.elements.searchResults.classList.remove('active');
            this.renderFunctionList();
            return;
        }

        this.showSearchResults(query);
        this.renderFunctionList(query);
    }

    /**
     * 검색 결과 표시
     */
    showSearchResults(query) {
        const results = window.searchEngine.search(query, { limit: 10 });

        if (results.length === 0) {
            this.elements.searchResults.innerHTML = '<div class="no-results">검색 결과가 없습니다</div>';
        } else {
            let html = '';
            results.forEach(func => {
                const desc = func.description
                    ? (func.description.length > 50 ? func.description.substring(0, 50) + '...' : func.description)
                    : '';
                html += `<div class="search-result-item" data-function="${func.name}">
                    <div class="func-name">${func.name}</div>
                    <div class="func-desc">${desc}</div>
                </div>`;
            });
            this.elements.searchResults.innerHTML = html;
        }

        this.elements.searchResults.classList.add('active');
    }

    /**
     * 해시 변경 처리
     */
    handleHashChange() {
        const hash = window.location.hash.substring(1);
        if (hash) {
            const funcName = decodeURIComponent(hash);
            this.renderFunctionDetail(funcName);

            // 모바일에서 사이드바 닫기
            this.elements.sidebar.classList.remove('open');
            this.elements.sidebarOverlay.classList.remove('active');
        } else {
            this.showWelcome();
        }
    }

    /**
     * HTML 이스케이프
     */
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * 오류 표시
     */
    showError(message) {
        this.elements.welcomeView.innerHTML = `
            <div class="welcome-content">
                <h1>오류 발생</h1>
                <p class="welcome-desc">${message}</p>
            </div>
        `;
    }
}

// 앱 시작
document.addEventListener('DOMContentLoaded', () => {
    const app = new ApiDocApp();
    app.init();
});
