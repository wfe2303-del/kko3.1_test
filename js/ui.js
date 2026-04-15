(function(){
  var utils = window.KakaoCheckUtils;
  var modalState = null;
  var loadingDepth = 0;

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
      if(event.target === els.overlay){
        closeModal();
      }
    });
    document.addEventListener('keydown', function(event){
      if(event.key === 'Escape' && modalState){
        closeModal();
      }
    });
  }

  function openModal(title, subtitle){
    var els;
    bindModalShell();
    els = modalEls();
    els.title.textContent = title || '';
    els.subtitle.textContent = subtitle || '';
    els.body.innerHTML = '';
    els.overlay.classList.remove('hidden');
    els.overlay.setAttribute('aria-hidden', 'false');
    return els.body;
  }

  function closeModal(){
    var els = modalEls();
    if(!els.overlay) return;
    els.overlay.classList.add('hidden');
    els.overlay.setAttribute('aria-hidden', 'true');
    els.body.innerHTML = '';
    modalState = null;
  }

  function openCustomModal(options){
    var body = openModal(options.title || '', options.subtitle || '');
    modalState = {
      type: options.type || 'custom',
      panelId: options.panelId || null
    };
    if(typeof options.render === 'function'){
      options.render(body);
    }
  }

  function getModalState(){
    return modalState;
  }

  function showLoading(title, message){
    var overlay = document.getElementById('loadingOverlay');
    var titleEl = document.getElementById('loadingTitle');
    var textEl = document.getElementById('loadingText');

    loadingDepth += 1;
    if(!overlay) return;

    titleEl.textContent = title || '처리 중입니다';
    textEl.textContent = message || '잠시만 기다려 주세요.';
    overlay.classList.remove('hidden');
    overlay.setAttribute('aria-hidden', 'false');
  }

  function hideLoading(force){
    var overlay = document.getElementById('loadingOverlay');
    if(force){
      loadingDepth = 0;
    } else {
      loadingDepth = Math.max(0, loadingDepth - 1);
    }

    if(!overlay || loadingDepth > 0) return;
    overlay.classList.add('hidden');
    overlay.setAttribute('aria-hidden', 'true');
  }

  function setAuthState(state){
    var isLoggedIn = !!(state && state.user && state.accessToken);
    var loginScreen = document.getElementById('loginScreen');
    var appShell = document.getElementById('appShell');
    var logoutBtn = document.getElementById('logoutBtn');
    var topbarUser = document.getElementById('topbarUser');

    if(loginScreen) loginScreen.classList.toggle('hidden', isLoggedIn);
    if(appShell) appShell.classList.toggle('hidden', !isLoggedIn);
    if(logoutBtn) logoutBtn.classList.toggle('hidden', !isLoggedIn);
    if(topbarUser){
      topbarUser.textContent = isLoggedIn ? formatUserLabel(state.user) : '';
      topbarUser.classList.toggle('hidden', !isLoggedIn);
    }
  }

  function setLoginReady(isReady){
    var loginBtn = document.getElementById('loginBtn');
    var usernameInput = document.getElementById('loginUsername');
    var passwordInput = document.getElementById('loginPassword');

    if(loginBtn){
      loginBtn.disabled = !isReady;
      loginBtn.textContent = isReady ? '로그인' : '로그인 중...';
    }
    if(usernameInput) usernameInput.disabled = !isReady;
    if(passwordInput) passwordInput.disabled = !isReady;
  }

  function setAuthError(message){
    var authErr = document.getElementById('authErr');
    if(authErr){
      authErr.textContent = message || '';
    }
  }

  function setAppNotice(message, kind){
    var notice = document.getElementById('appNotice');
    if(!notice) return;

    if(!message){
      notice.className = 'card app-notice hidden';
      notice.textContent = '';
      return;
    }

    notice.className = 'card app-notice ' + (kind || 'warn');
    notice.textContent = message;
  }

  function createPanelElement(panelId){
    var template = document.getElementById('panelTemplate');
    var node = template.content.firstElementChild.cloneNode(true);
    node.dataset.panelId = String(panelId);
    return node;
  }

  function renderSelectedSheet(panelEl, sheetTitle){
    var node = utils.qs('.js-selected-sheet', panelEl);
    var titleNode = utils.qs('.panel-title', panelEl);
    var label = sheetTitle || '선택되지 않음';

    if(node) node.textContent = label;
    if(titleNode) titleNode.textContent = sheetTitle || '시트를 선택해주세요';
  }

  function renderFileSummary(panelEl, files){
    var summary = utils.qs('.js-file-summary', panelEl);
    if(!summary) return;

    if(!files.length){
      summary.textContent = '파일 없음';
      return;
    }

    summary.textContent = files.length === 1 ? files[0].name : (files.length + '개 파일');
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
    if(errorBox){
      errorBox.textContent = message || '';
    }
  }

  function setPanelActionState(panelEl, state){
    var manualBtn = utils.qs('.js-manual-review-btn', panelEl);
    var storageBtn = utils.qs('.js-storage-data-btn', panelEl);

    if(manualBtn){
      manualBtn.disabled = !state.manualEnabled;
      manualBtn.textContent = state.manualLabel || '미매칭 수동 처리';
    }

    if(storageBtn){
      storageBtn.disabled = !state.storageEnabled;
    }
  }

  function renderResults(panelEl, sections){
    var root = utils.qs('.js-panel-results', panelEl);
    var summarySection = null;
    var tableSections = [];

    if(!root) return;
    root.innerHTML = '';

    if(!sections || !sections.length){
      appendEmptyState(root, '결과가 없습니다.');
      return;
    }

    sections.forEach(function(section){
      if(section && section.summary){
        summarySection = section;
      } else if(section){
        tableSections.push(section);
      }
    });

    if(tableSections.length){
      var buttonGrid = document.createElement('div');
      buttonGrid.className = 'result-button-grid';

      tableSections.forEach(function(section){
        var button = document.createElement('button');
        var labelWrap = document.createElement('span');
        var badge = document.createElement('span');

        button.type = 'button';
        button.className = 'result-open-btn';
        labelWrap.textContent = section.title;
        badge.className = 'result-count-badge';
        badge.textContent = String(section.rows ? section.rows.length : 0);

        button.appendChild(labelWrap);
        button.appendChild(badge);
        button.addEventListener('click', function(){
          openSectionResultModal(section);
        });
        buttonGrid.appendChild(button);
      });

      root.appendChild(buttonGrid);
    }

    if(summarySection){
      root.appendChild(buildMetricGrid(summarySection.summary || []));
    }
  }

  function openSectionResultModal(section){
    openCustomModal({
      type: 'result-section',
      title: section.title,
      subtitle: section.rows && section.rows.length ? (section.rows.length + '건') : '데이터 없음',
      render: function(body){
        if(section.groups && section.groups.length){
          renderGroupedSection(body, section);
          return;
        }

        if(!section.rows || !section.rows.length){
          appendEmptyState(body, '데이터가 없습니다.');
          return;
        }

        body.appendChild(buildTableWrap(section.headers || [], section.rows || []));
      }
    });
  }

  function renderGroupedSection(root, section){
    var groups = section.groups || [];
    var activeKey = groups[0] ? groups[0].key : '';
    var filterRow = document.createElement('div');
    var content = document.createElement('div');

    filterRow.className = 'result-filter-row';

    function draw(){
      var current = groups.find(function(group){
        return group.key === activeKey;
      }) || groups[0];

      filterRow.innerHTML = '';
      content.innerHTML = '';

      groups.forEach(function(group){
        var button = document.createElement('button');
        button.type = 'button';
        button.className = 'result-filter-btn' + (group.key === activeKey ? ' active' : '');
        button.textContent = group.label + ' ' + (group.rows ? group.rows.length : 0);
        button.addEventListener('click', function(){
          activeKey = group.key;
          draw();
        });
        filterRow.appendChild(button);
      });

      if(!current || !current.rows || !current.rows.length){
        appendEmptyState(content, '데이터가 없습니다.');
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
    var table = document.createElement('table');
    var thead = document.createElement('thead');
    var headerRow = document.createElement('tr');
    var tbody = document.createElement('tbody');

    wrap.className = 'table-wrap';

    headers.forEach(function(header){
      var th = document.createElement('th');
      th.textContent = header;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    rows.forEach(function(row){
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

  function buildMetricGrid(items){
    var grid = document.createElement('div');
    grid.className = 'metric-grid';

    items.forEach(function(item){
      var card = document.createElement('div');
      var label = document.createElement('span');
      var value = document.createElement('strong');

      card.className = 'metric-card';
      label.className = 'metric-label';
      value.className = 'metric-value';
      label.textContent = item.label || '';
      value.textContent = item.value == null ? '' : String(item.value);

      card.appendChild(label);
      card.appendChild(value);
      grid.appendChild(card);
    });

    return grid;
  }

  function appendEmptyState(root, message){
    var empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.textContent = message || '데이터가 없습니다.';
    root.appendChild(empty);
  }

  function openSheetPickerModal(options){
    openCustomModal({
      type: 'sheet-picker',
      panelId: options.panelId,
      title: options.title || '시트 선택',
      subtitle: options.subtitle || '',
      render: function(body){
        var search = document.createElement('input');
        var list = document.createElement('div');

        search.className = 'input modal-search';
        search.type = 'text';
        search.placeholder = '시트 이름 검색';
        search.value = options.initialQuery || '';

        list.className = 'modal-list';

        function renderList(){
          var query = String(search.value || '').trim().toLowerCase();
          var filtered = (options.titles || []).filter(function(title){
            return !query || title.toLowerCase().indexOf(query) >= 0;
          });

          list.innerHTML = '';
          if(!filtered.length){
            appendEmptyState(list, options.emptyText || '검색 결과가 없습니다.');
            return;
          }

          filtered.forEach(function(title){
            var button = document.createElement('button');
            button.type = 'button';
            button.className = 'modal-list-btn' + (title === options.selectedTitle ? ' active' : '');
            button.textContent = title;
            button.addEventListener('click', function(){
              options.onSelect(title);
              closeModal();
            });
            list.appendChild(button);
          });
        }

        search.addEventListener('input', renderList);
        body.appendChild(search);
        body.appendChild(list);
        renderList();
      }
    });
  }

  function openFileManagerModal(options){
    openCustomModal({
      type: 'file-manager',
      panelId: options.panelId,
      title: options.title || '로그 파일 관리',
      subtitle: options.subtitle || '',
      render: function(body){
        var actions = document.createElement('div');
        var addBtn = document.createElement('button');
        var clearBtn = document.createElement('button');
        var list = document.createElement('div');

        actions.className = 'file-manager-actions';
        addBtn.type = 'button';
        addBtn.className = 'btn btn-primary';
        addBtn.textContent = '파일 추가';
        addBtn.addEventListener('click', function(){
          options.onAddRequest();
        });

        clearBtn.type = 'button';
        clearBtn.className = 'btn';
        clearBtn.textContent = '전체 비우기';
        clearBtn.addEventListener('click', function(){
          options.onClearAll();
          openFileManagerModal(options.getFreshOptions());
        });

        actions.appendChild(addBtn);
        actions.appendChild(clearBtn);
        body.appendChild(actions);

        list.className = 'modal-list';
        if(!options.files || !options.files.length){
          appendEmptyState(list, '추가된 로그 파일이 없습니다.');
          body.appendChild(list);
          return;
        }

        options.files.forEach(function(file, index){
          var item = document.createElement('div');
          var main = document.createElement('div');
          var fileName = document.createElement('div');
          var meta = document.createElement('div');
          var removeBtn = document.createElement('button');

          item.className = 'file-item';
          main.className = 'file-item-main';
          fileName.className = 'file-name';
          meta.className = 'file-meta';
          fileName.textContent = file.name;
          meta.textContent = formatBytes(file.size) + ' · ' + new Date(file.lastModified || Date.now()).toLocaleString();

          removeBtn.type = 'button';
          removeBtn.className = 'btn btn-danger';
          removeBtn.textContent = '제거';
          removeBtn.addEventListener('click', function(){
            options.onRemoveIndex(index);
            openFileManagerModal(options.getFreshOptions());
          });

          main.appendChild(fileName);
          main.appendChild(meta);
          item.appendChild(main);
          item.appendChild(removeBtn);
          list.appendChild(item);
        });

        body.appendChild(list);
      }
    });
  }

  function formatBytes(bytes){
    var units = ['B', 'KB', 'MB', 'GB'];
    var value = Number(bytes || 0);
    var unitIndex = 0;

    while(value >= 1024 && unitIndex < units.length - 1){
      value /= 1024;
      unitIndex += 1;
    }

    return (unitIndex === 0 ? value : value.toFixed(1)) + ' ' + units[unitIndex];
  }

  function formatUserLabel(user){
    var displayName = String(user && (user.displayName || user.name) || '').trim();
    var username = String(user && (user.username || user.email) || '').trim();
    if(displayName && username && displayName !== username){
      return displayName + ' (' + username + ')';
    }
    return displayName || username || '';
  }

  window.KakaoCheckUI = {
    appendEmptyState: appendEmptyState,
    buildMetricGrid: buildMetricGrid,
    buildTableWrap: buildTableWrap,
    closeModal: closeModal,
    createPanelElement: createPanelElement,
    getModalState: getModalState,
    openCustomModal: openCustomModal,
    openFileManagerModal: openFileManagerModal,
    openSheetPickerModal: openSheetPickerModal,
    renderFileSummary: renderFileSummary,
    renderResults: renderResults,
    renderSelectedSheet: renderSelectedSheet,
    hideLoading: hideLoading,
    showLoading: showLoading,
    setAppNotice: setAppNotice,
    setAuthError: setAuthError,
    setAuthState: setAuthState,
    setLoginReady: setLoginReady,
    setPanelActionState: setPanelActionState,
    setPanelError: setPanelError,
    setPanelStatus: setPanelStatus
  };
})();
