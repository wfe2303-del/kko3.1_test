(function(){
  function qs(selector, root){ return (root || document).querySelector(selector); }
  function qsa(selector, root){ return Array.from((root || document).querySelectorAll(selector)); }
  function digitsOnly(value){ return String(value || '').replace(/\D+/g, ''); }
  function last4Digits(value){ return digitsOnly(value).slice(-4); }
  function normalizeName(value){
    var text = String(value || '').trim();
    text = text.replace(/[\s]*[\(\[\{（【][^)\]\}）】]*[\)\]\}）】]\s*$/, '');
    text = text.replace(/[/_\-\s]+$/, '');
    text = text.replace(/\s+/g, ' ');
    return text;
  }
  function escapeHtml(value){
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
  function readFileAsText(file){
    return new Promise(function(resolve, reject){
      var reader = new FileReader();
      reader.onload = function(event){ resolve(String((event.target && event.target.result) || '')); };
      reader.onerror = function(){ reject(new Error('파일을 읽는 중 오류가 발생했습니다.')); };
      reader.readAsText(file, 'utf-8');
    });
  }
  function csvToRows(text){
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    for(var i = 0; i < text.length; i += 1){
      var ch = text[i];
      var next = text[i + 1];
      if(ch === '"'){
        if(inQuotes && next === '"'){
          field += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if(ch === ',' && !inQuotes){
        row.push(field);
        field = '';
      } else if((ch === '\n' || ch === '\r') && !inQuotes){
        if(ch === '\r' && next === '\n') i += 1;
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
      } else {
        field += ch;
      }
    }
    if(field.length > 0 || row.length > 0){
      row.push(field);
      rows.push(row);
    }
    return rows;
  }
  function statusClass(kind){
    if(kind === 'ok') return 'status-pill ok';
    if(kind === 'bad') return 'status-pill bad';
    if(kind === 'warn') return 'status-pill warn';
    return 'status-pill';
  }
  window.KakaoCheckUtils = {
    qs: qs,
    qsa: qsa,
    digitsOnly: digitsOnly,
    last4Digits: last4Digits,
    normalizeName: normalizeName,
    escapeHtml: escapeHtml,
    readFileAsText: readFileAsText,
    csvToRows: csvToRows,
    statusClass: statusClass
  };
})();
