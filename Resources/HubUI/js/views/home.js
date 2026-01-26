// Resources/HubUI/js/views/home.js
import { clear, div } from '../core/dom.js';

const MULTI_MODE_KEY = 'kky.hub.multiMode';

export function renderHome(root) {
    const target = root || document.getElementById('view-root') || document.getElementById('app');
    clear(target);

    const view = div('home-choice');
    const hero = div('home-choice-hero');
    hero.innerHTML = `
        <p class="home-choice-kicker">KKY Tool Hub</p>
        <h2>검토 방식을 선택하세요</h2>
        <p>활성 문서 기반 검토 또는 다중 RVT 배치 검토를 시작할 수 있습니다.</p>`;

    const grid = div('home-choice-grid');
    grid.append(
        buildCard(
            '활성 문서 검토',
            '현재 열려있는 Revit 문서를 대상으로 빠르게 검토를 수행합니다.',
            'active-menu',
            [
                '중복 객체 검토: 현재 열린 문서에서 중복 요소/패밀리 점검',
                '복합 패밀리 공유파라미터 추가 및 연동: 공유 파라미터 추가/연동 수행'
            ]
        ),
        buildCard(
            '납품시 BQC 검토',
            '납품 기준에 맞춰 유틸리티 도구를 실행합니다.',
            'multi',
            [
                '파라미터 값 연속성 검토: 연결된 객체들의 파라미터 값 연속성 검토',
                '공유파라미터 GUID 검토: 프로젝트/패밀리 내 공유 파라미터 GUID 검토',
                '패밀리 공유파라미터 연동 검토: 복합 패밀리 연동 상태 점검',
                'Point 추출: Project/Survey 포인트 좌표 추출'
            ],
            'bqc'
        ),
        buildCard(
            '유틸리티',
            '납품시 BQC 검토 유틸리티 도구 모음',
            'multi',
            [
                'PMS 검토: Segment ↔ PMS 매핑 및 사이즈 검토',
                'GUID/연동/Point 추출/Project Parameter 추가',
            ],
            'utility',
            'utilities'
        )
    );

    view.append(hero, grid);
    target.append(view);

    function buildCard(title, desc, hash, items, multiMode, anchorId) {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'home-choice-card';
        const listHtml = Array.isArray(items) && items.length
          ? `<ul class="home-choice-list">${items.map((item) => `<li>${item}</li>`).join('')}</ul>`
          : '';
        card.innerHTML = `
            <div class="home-choice-card__body">
              <div>
                <h3>${title}</h3>
                <p>${desc}</p>
                ${listHtml}
              </div>
              <span class="home-choice-card__icon">→</span>
            </div>
            <span class="home-choice-cta btn btn--primary">바로가기</span>`;
        card.addEventListener('click', () => {
            if (multiMode) setMultiMode(multiMode);
            location.hash = `#${hash}`;
            if (anchorId) {
                setTimeout(() => {
                    const el = document.getElementById(anchorId);
                    if (el && el.scrollIntoView) {
                        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                }, 240);
            }
        });
        return card;
    }

    function setMultiMode(mode) {
        try {
            localStorage.setItem(MULTI_MODE_KEY, mode);
        } catch {
        }
    }
}
