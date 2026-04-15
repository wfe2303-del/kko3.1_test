(function(){
  function qs(selector, root){
    return (root || document).querySelector(selector);
  }

  function qsa(selector, root){
    return Array.from((root || document).querySelectorAll(selector));
  }

  function digitsOnly(value){
    return String(value || '').replace(/\D+/g, '');
  }

  function last4Digits(value){
    return digitsOnly(value).slice(-4);
  }

  function normalizeName(value){
    var text = String(value || '').trim();
    text = text.replace(/[\s]*[\(\[\{][^)\]\}]*[\)\]\}]\s*$/, '');
    text = text.replace(/[/_\-\s]+$/, '');
    text = text.replace(/\s+/g, ' ');
    return text;
  }

  function readFileAsText(file){
    return new Promise(function(resolve, reject){
      var reader = new FileReader();
      reader.onload = function(event){
        resolve(String((event.target && event.target.result) || ''));
      };
      reader.onerror = function(){
        reject(new Error('파일을 읽는 중 오류가 발생했습니다.'));
      };
      reader.readAsText(file, 'utf-8');
    });
  }

  function csvToRows(text){
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;

    for(var index = 0; index < text.length; index += 1){
      var ch = text[index];
      var next = text[index + 1];

      if(ch === '"'){
        if(inQuotes && next === '"'){
          field += '"';
          index += 1;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }

      if(ch === ',' && !inQuotes){
        row.push(field);
        field = '';
        continue;
      }

      if((ch === '\n' || ch === '\r') && !inQuotes){
        if(ch === '\r' && next === '\n'){
          index += 1;
        }
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        continue;
      }

      field += ch;
    }

    if(field.length || row.length){
      row.push(field);
      rows.push(row);
    }

    return rows;
  }

  function statusClass(kind){
    if(kind === 'ok') return 'status-pill ok';
    if(kind === 'warn') return 'status-pill warn';
    if(kind === 'bad') return 'status-pill bad';
    return 'status-pill';
  }

  function formatPersonLabel(name, phone4){
    var displayName = String(name || '').trim();
    var digits = last4Digits(phone4);
    return digits ? displayName + '/' + digits : displayName;
  }

  window.KakaoCheckUtils = {
    qs: qs,
    qsa: qsa,
    digitsOnly: digitsOnly,
    last4Digits: last4Digits,
    normalizeName: normalizeName,
    readFileAsText: readFileAsText,
    csvToRows: csvToRows,
    statusClass: statusClass,
    formatPersonLabel: formatPersonLabel
  };
})();
