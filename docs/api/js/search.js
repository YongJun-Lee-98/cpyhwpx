/**
 * cpyhwpx API 검색 엔진
 */

class SearchEngine {
    constructor() {
        this.data = null;
        this.index = new Map();
    }

    /**
     * 데이터 설정 및 인덱스 구축
     */
    setData(data) {
        this.data = data;
        this.buildIndex();
    }

    /**
     * 역인덱스 구축
     */
    buildIndex() {
        this.index.clear();

        this.data.functions.forEach(func => {
            // 함수명 토큰화
            const nameTokens = this.tokenize(func.name);
            nameTokens.forEach(token => {
                if (!this.index.has(token)) {
                    this.index.set(token, new Set());
                }
                this.index.get(token).add(func.name);
            });

            // 설명 토큰화
            if (func.description) {
                const descTokens = this.tokenize(func.description);
                descTokens.forEach(token => {
                    if (!this.index.has(token)) {
                        this.index.set(token, new Set());
                    }
                    this.index.get(token).add(func.name);
                });
            }
        });
    }

    /**
     * 텍스트 토큰화 (한글, 영문 모두 처리)
     */
    tokenize(text) {
        if (!text) return [];

        return text.toLowerCase()
            .replace(/[_\-\.]/g, ' ')  // 구분자를 공백으로
            .split(/\s+/)
            .filter(t => t.length > 1)
            .map(t => t.trim());
    }

    /**
     * 검색 실행
     */
    search(query, options = {}) {
        if (!this.data || !query || query.trim().length < 1) {
            return [];
        }

        const {
            category = 'all',
            limit = 50,
            includeDescription = true
        } = options;

        const q = query.toLowerCase().trim();
        let results = this.data.functions;

        // 카테고리 필터
        if (category !== 'all') {
            results = results.filter(f => f.category === category);
        }

        // 검색 점수 계산
        results = results
            .map(func => ({
                ...func,
                score: this.calculateScore(func, q, includeDescription)
            }))
            .filter(f => f.score > 0)
            .sort((a, b) => b.score - a.score)
            .slice(0, limit);

        return results;
    }

    /**
     * 검색 점수 계산
     */
    calculateScore(func, query, includeDescription) {
        let score = 0;
        const name = func.name.toLowerCase();
        const q = query.toLowerCase();

        // 함수명 정확 일치
        if (name === q) {
            score += 100;
        }
        // 함수명 시작
        else if (name.startsWith(q)) {
            score += 50;
        }
        // 함수명 포함
        else if (name.includes(q)) {
            score += 20;
        }

        // 언더스코어로 구분된 부분 매칭
        const nameParts = name.split('_');
        for (const part of nameParts) {
            if (part === q) {
                score += 30;
                break;
            } else if (part.startsWith(q)) {
                score += 15;
                break;
            }
        }

        // 설명 매칭
        if (includeDescription && func.description) {
            const desc = func.description.toLowerCase();
            if (desc.includes(q)) {
                score += 10;

                // 설명 시작 부분에 있으면 추가 점수
                if (desc.startsWith(q)) {
                    score += 5;
                }
            }
        }

        return score;
    }

    /**
     * 자동완성 제안
     */
    suggest(query, limit = 5) {
        if (!query || query.length < 1) {
            return [];
        }

        const q = query.toLowerCase();
        const suggestions = [];

        // 함수명에서 접두사 매칭
        for (const func of this.data.functions) {
            const name = func.name.toLowerCase();
            if (name.startsWith(q) || name.includes('_' + q)) {
                suggestions.push({
                    name: func.name,
                    description: func.description
                });
            }
            if (suggestions.length >= limit) break;
        }

        return suggestions;
    }

    /**
     * 카테고리 필터링
     */
    filterByCategory(category) {
        if (!this.data) return [];

        if (category === 'all') {
            return [...this.data.functions];
        }

        return this.data.functions.filter(f => f.category === category);
    }

    /**
     * 연관 함수 찾기
     */
    findRelated(func, limit = 5) {
        if (!this.data || !func) return [];

        const related = [];
        const funcName = func.name.toLowerCase();

        // 1. @requires, @before로 연결된 함수
        if (func.requires) {
            func.requires.forEach(req => {
                const found = this.data.functions.find(f =>
                    f.name.toLowerCase() === req.toLowerCase()
                );
                if (found) {
                    related.push({
                        ...found,
                        relationType: 'requires'
                    });
                }
            });
        }

        if (func.before) {
            func.before.forEach(before => {
                const found = this.data.functions.find(f =>
                    f.name.toLowerCase() === before.toLowerCase()
                );
                if (found) {
                    related.push({
                        ...found,
                        relationType: 'before'
                    });
                }
            });
        }

        // 2. 같은 카테고리 함수
        const sameCategory = this.data.functions
            .filter(f => f.category === func.category && f.name !== func.name)
            .slice(0, limit);

        sameCategory.forEach(f => {
            if (!related.find(r => r.name === f.name)) {
                related.push({
                    ...f,
                    relationType: 'category'
                });
            }
        });

        return related.slice(0, limit + (func.requires?.length || 0) + (func.before?.length || 0));
    }
}

// 전역 인스턴스
window.searchEngine = new SearchEngine();
