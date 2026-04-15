(function(){
  var utils = window.KakaoCheckUtils;
  var modalState = null;

  function modalEls(){
    return {
      overlay: document.getElementById('modalOverlay'),
      title: document.getElementById('modalTitle'),
      subtitle: document.getElementById('modalSubtitle'),
      body: document.getElementById('modalBody'),
      closeBtn: document.getElementById('modalCloseBtn')
    };
  }

  function bindModalShell(){
    var els = modalEls();
    if(!els.overlay || els.overlay.dataset.bound === '1') return;
    els.overlay.dataset.bound = '1';
    els.closeBtn.addEventListener('click', closeModal);
    els.overlay.addEventListener('click', function(event){
      if(event.target === els.overlay) closeModal();
    });
    document.addEventListener('keydown', function(event){
      if(event.key === 'Escape' && modalState) closeModal();
    });
  }

  function openModal(title, subtitle){
    bindModalShell();
    var els = modalEls();
    els.title.textContent = title || '';
    els.subtitle.textContent = subtitle || '';
    els.body.innerHTML = '';
    els.overlay.classList.remove('hidden');
    els.overlay.setAttribute('aria-hidden', 'false');
  }

  function closeModal(){
    var els = modalEls();
    if(!els.overlay) return;
    els.overlay.classList.add('hidden');
    els.overlay.setAttribute('aria-hidden', 'true');
    els.body.innerHTML = '';
    modalState = null;
  }

  function getModalState(){
    return modalState;
  }

  function setAuthState(state){
    var loginScreen = document.getElementById('loginScreen');
    var appShell = document.getElementById('appShell');
    var logoutBtn = document.getElementById('logoutBtn');
    var topbarUser = document.getElementById('topbarUser');
    var isLoggedIn = !!(state && state.user && state.user.email && state.accessToken);

    if(loginScreen) loginScreen.classList.toggle('hidden', isLoggedIn);
    if(appShell) appShell.classList.toggle('hidden', !isLoggedIn);
    if(logoutBtn) logoutBtn.classList.toggle('hidden', !isLoggedIn);
    if(topbarUser){
      topbarUser.textContent = isLoggedIn ? state.user.email : '';
      topbarUser.classList.toggle('hidden', !isLoggedIn);
    }
  }

  function setLoginReady(isReady){
    var loginBtn = document.getElementById('loginBtn');
    if(!loginBtn) return;
    loginBtn.disabled = !isReady;
    loginBtn.textContent = isReady ? 'Google 계정으로 로그인' : '로그인 준비 중...';
  }

  function setTabState(message, kind){
    var tabState = document.getElementById('tabState');
    if(!tabState) return;
    tabState.className = utils.statusClass(kind || '');
    tabState.textContent = message || '';
  }

  function setAuthError(message){
    var authErr = document.getElementById('authErr');
    if(authErr) authErr.textContent = message || '';
  }

  function createPanelElement(panelId){
    var template = document.getElementById('panelTemplate');
    var node = template.content.firstElementChild.cloneNode(true);
    node.dataset.panelId = String(panelId);
    utils.qs('.panel-title', node).textContent = '탭 선택 전';
    return node;
  }

  function renderSelectedSheet(panelEl, sheetTitle){
    var node = utils.qs('.js-selected-sheet', panelEl);
    var titleNode = utils.qs('.panel-title', panelEl);
    var label = sheetTitle || '탭 미선택';
    if(node) node.textContent = label;
    if(titleNode) titleNode.textContent = sheetTitle || '탭 선택 전';
  }

  function renderFileSummary(panelEl, files){
    var summary = utils.qs('.js-file-summary', panelEl);
    if(!summary) return;
    if(!files.length){
      summary.textContent = '파일 없음';
      return;
    }
    if(files.length === 1){
      summary.textContent = files[0].name;
      return;
    }
    summary.textContent = files.length + '개 파일';
  }

  function setPanelStatus(panelEl, message, kind){
    var status = utils.qs('.js-panel-status', panelEl);
    var meta = utils.qs('.panel-meta', panelEl);
    if(!status || !meta) return;
    if(!message){
      meta.classList.add('hidden');
      status.className = 'status-pill js-panel-status';
      status.textContent = '';
      return;
    }
    meta.classList.remove('hidden');
    status.className = utils.statusClass(kind || '');
    status.textContent = message;
  }

  function setPanelError(panelEl, message){
    var errorBox = utils.qs('.js-panel-error', panelEl);
    if(errorBox) errorBox.textContent = message || '';
  }

  function renderResults(panelEl, sections){
    var root = utils.qs('.js-panel-results', panelEl);
    if(!root) return;
    root.innerHTML = '';
    if(!sections.length){
      root.innerHTML = '<div class="empty-state">결과가 없습니다.</div>';
      return;
    }

    var summarySection = null;
    var listSections = [];
    sections.forEach(function(section){
      if(section.title === '로그 요약') summarySection = section;
      else listSections.push(section);
    });

    if(listSections.length){
      var buttonGrid = document.createElement('div');
      buttonGrid.className = 'result-button-grid';
      listSections.forEach(function(section){
        var button = document.createElement('button');
        button.type = 'button';
        button.className = 'result-open-btn';
        var titleWrap = document.createElement('span');
        titleWrap.textContent = section.title;
        var badge = document.createElement('span');
        badge.className = 'result-count-badge';
        badge.textContent = String(section.rows ? section.rows.length : 0);
        button.appendChild(titleWrap);
        button.appendChild(badge);
        button.addEventListener('click', function(){
          openSectionResultModal(section);
        });
        buttonGrid.appendChild(button);
      });
      root.appendChild(buttonGrid);
    }

    if(summarySection && summarySection.rows && summarySection.rows[0]){
      var summaryCard = document.createElement('div');
      summaryCard.className = 'result-summary-card';
      var title = document.createElement('h3');
      title.className = 'result-summary-title';
      title.textContent = summarySection.title;
      summaryCard.appendChild(title);
      var summaryGrid = document.createElement('div');
      summaryGrid.className = 'result-summary-grid';
      summarySection.headers.forEach(function(header, index){
        var stat = document.createElement('div');
        stat.className = 'result-stat';
        var label = document.createElement('span');
        label.className = 'result-stat-label';
        label.textContent = header;
        var value = document.createElement('strong');
        value.className = 'result-stat-value';
        value.textContent = summarySection.rows[0][index] == null ? '' : String(summarySection.rows[0][index]);
        stat.appendChild(label);
        stat.appendChild(value);
        summaryGrid.appendChild(stat);
      });
      summaryCard.appendChild(summaryGrid);
      root.appendChild(summaryCard);
    }
  }

  function openSectionResultModal(section){
    openModal(section.title, (section.rows && section.rows.length ? section.rows.length + '건' : '데이터 없음'));
    modalState = { type: 'result-section', title: section.title };
    var els = modalEls();

    if(section.groups && section.groups.length){
      renderGroupedSectionModal(els.body, section);
      return;
    }

    if(!section.rows || !section.rows.length){
      var empty = document.createElement('div');
      empty.className = 'empty-state';
      empty.textContent = '데이터 없음';
      els.body.appendChild(empty);
      return;
    }

    els.body.appendChild(buildTableWrap(section.headers || [], section.rows || []));
  }

  function renderGroupedSectionModal(root, section){
    var groups = section.groups || [];
    var activeKey = groups[0] ? groups[0].key : '';
    var filterRow = document.createElement('div');
    filterRow.className = 'result-filter-row';
    var content = document.createElement('div');

    function draw(){
      filterRow.innerHTML = '';
      content.innerHTML = '';
      groups.forEach(function(group){
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'result-filter-btn' + (group.key === activeKey ? ' active' : '');
        btn.textContent = group.label + ' ' + (group.rows ? group.rows.length : 0);
        btn.addEventListener('click', function(){
          activeKey = group.key;
          draw();
        });
        filterRow.appendChild(btn);
      });

      var current = groups.find(function(group){ return group.key === activeKey; }) || groups[0];
      if(!current || !current.rows || !current.rows.length){
        var empty = document.createElement('div');
        empty.className = 'empty-state';
        empty.textContent = '데이터 없음';
        content.appendChild(empty);
        return;
      }
      content.appendChild(buildTableWrap(section.headers || [], current.rows || []));
    }

    root.appendChild(filterRow);
    root.appendChild(content);
    draw();
  }

  function buildTableWrap(headers, rows){
    var wrap = document.createElement('div');
    wrap.className = 'table-wrap';
    var table = document.createElement('table');
    var thead = document.createElement('thead');
    var headerRow = document.createElement('tr');
    (headers || []).forEach(function(header){
      var th = document.createElement('th');
      th.textContent = header;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    var tbody = document.createElement('tbody');
    (rows || []).forEach(function(row){
      var tr = document.createElement('tr');
      row.forEach(function(cell){
        var td = document.createElement('td');
        td.textContent = cell == null ? '' : String(cell);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
    return wrap;
  }

  function openSheetPickerModal(options){
    openModal(options.title || '탭 선택', options.subtitle || '');
    modalState = { type: 'sheet-picker', panelId: options.panelId };
    var els = modalEls();
    var search = document.createElement('input');
    search.className = 'input modal-search';
    search.type = 'text';
    search.placeholder = '탭 이름 검색';
    search.value = options.initialQuery || '';
    var list = document.createElement('div');
    list.className = 'modal-list';

    function renderList(){
      var query = String(search.value || '').trim().toLowerCase();
      list.innerHTML = '';
      var filtered = (options.titles || []).filter(function(title){
        return !query || title.toLowerCase().indexOf(query) >= 0;
      });
      if(!filtered.length){
        var empty = document.createElement('div');
        empty.className = 'file-empty';
        empty.textContent = '검색 결과가 없습니다.';
        list.appendChild(empty);
        return;
      }
      filtered.forEach(function(titleText){
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'modal-list-btn' + (titleText === options.selectedTitle ? ' active' : '');
        btn.textContent = titleText;
        btn.addEventListener('click', function(){
          options.onSelect(titleText);
          closeModal();
        });
        list.appendChild(btn);
      });
    }

    search.addEventListener('input', renderList);
    els.body.appendChild(search);
    els.body.appendChild(list);
    renderList();
  }

  function openFileManagerModal(options){
    openModal(options.title || '로그 파일 관리', options.subtitle || '');
    modalState = { type: 'file-manager', panelId: options.panelId };
    var els = modalEls();
    var actions = document.createElement('div');
    actions.className = 'file-manager-actions';

    var addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'btn btn-primary';
    addBtn.textContent = '파일 추가';
    addBtn.addEventListener('click', function(){ options.onAddRequest(); });
    actions.appendChild(addBtn);

    var clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'btn';
    clearBtn.textContent = '전체 비우기';
    clearBtn.addEventListener('click', function(){
      options.onClearAll();
      openFileManagerModal(options.getFreshOptions());
    });
    actions.appendChild(clearBtn);

    els.body.appendChild(actions);
    var list = document.createElement('div');
    list.className = 'modal-list';

    if(!options.files || !options.files.length){
      var empty = document.createElement('div');
      empty.className = 'file-empty';
      empty.textContent = '추가된 로그 파일이 없습니다.';
      list.appendChild(empty);
    } else {
      options.files.forEach(function(file, index){
        var item = document.createElement('div');
        item.className = 'file-item';
        var main = document.createElement('div');
        main.className = 'file-item-main';
        var fileName = document.createElement('div');
        fileName.className = 'file-name';
        fileName.textContent = file.name;
        var meta = document.createElement('div');
        meta.className = 'file-meta';
        meta.textContent = formatBytes(file.size) + ' · ' + new Date(file.lastModified || Date.now()).toLocaleString();
        main.appendChild(fileName);
        main.appendChild(meta);
        item.appendChild(main);
        var removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'btn btn-danger';
        removeBtn.textContent = '삭제';
        removeBtn.addEventListener('click', function(){
          options.onRemoveIndex(index);
          openFileManagerModal(options.getFreshOptions());
        });
        item.appendChild(removeBtn);
        list.appendChild(item);
      });
    }
    els.body.appendChild(list);
  }

  function formatBytes(bytes){
    if(!bytes) return '0 B';
    var units = ['B','KB','MB','GB'];
    var value = bytes;
    var unitIndex = 0;
    while(value >= 1024 && unitIndex < units.length - 1){
      value /= 1024;
      unitIndex += 1;
    }
    return (unitIndex === 0 ? value : value.toFixed(1)) + ' ' + units[unitIndex];
  }

  window.KakaoCheckUI = {
    closeModal: closeModal,
    getModalState: getModalState,
    setAuthState: setAuthState,
    setLoginReady: setLoginReady,
    setTabState: setTabState,
    setAuthError: setAuthError,
    createPanelElement: createPanelElement,
    renderSelectedSheet: renderSelectedSheet,
    renderFileSummary: renderFileSummary,
    setPanelStatus: setPanelStatus,
    setPanelError: setPanelError,
    renderResults: renderResults,
    openSheetPickerModal: openSheetPickerModal,
    openFileManagerModal: openFileManagerModal
  };
})();
