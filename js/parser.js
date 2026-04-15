(function(){
  var utils = window.KakaoCheckUtils;
  var JOIN_PAT = /(.*?)님이\s*(입장하셨습니다|들어왔습니다|입장했습니다|들어오셨습니다)[\s\.\!]*$/;
  var LEAVE_PAT = /(.*?)님이\s*(퇴장하셨습니다|나갔습니다|퇴장했습니다|나가셨습니다)[\s\.\!]*$/;
  var KICK_PAT = /(.*?)님을\s*(내보냈습니다|강퇴했습니다|추방했습니다)[\s\.\!]*$/;
  var NAME_ONLY_SENTINEL = '__NAME_ONLY__';

  function splitNickname(nick){
    var raw = String(nick || '').trim();
    if(!raw) return null;
    var idx = raw.search(/\d/);
    var namePart = idx >= 0 ? raw.slice(0, idx) : raw;
    var name = utils.normalizeName(namePart);
    if(!name) return { name: raw, digits: null };
    var digitsAll = raw.replace(/\D+/g, '');
    var digits = digitsAll.length >= 4 ? digitsAll.slice(-4) : null;
    return { name: name, digits: digits };
  }

  function parseChatText(text){
    var lines = String(text || '').split(/\r?\n/);
    var events = [];
    var joinedCount = 0;
    var leftCount = 0;

    lines.forEach(function(raw){
      var line = raw.trim();
      var match;
      if(!line) return;
      match = line.match(JOIN_PAT);
      if(match){
        events.push({ type: 'join', nick: match[1].trim(), raw: line });
        joinedCount += 1;
        return;
      }
      match = line.match(LEAVE_PAT);
      if(match){
        events.push({ type: 'leave', nick: match[1].trim(), raw: line });
        leftCount += 1;
        return;
      }
      match = line.match(KICK_PAT);
      if(match){
        events.push({ type: 'leave', nick: match[1].trim(), raw: line });
        leftCount += 1;
      }
    });

    return { events: events, joinedCount: joinedCount, leftCount: leftCount };
  }

  function extractLinesFromCsv(text){
    var rows = utils.csvToRows(text);
    var matched = [];
    var fallback = [];

    rows.forEach(function(row){
      if(!Array.isArray(row) || !row.length) return;
      row.forEach(function(cell, index){
        var value = String(cell || '').trim();
        if(!value) return;
        if(JOIN_PAT.test(value) || LEAVE_PAT.test(value) || KICK_PAT.test(value)){
          matched.push(value);
        }
        if(index === 2 && value) fallback.push(value);
      });
    });

    if(matched.length) return matched.join('\n');
    return fallback.join('\n');
  }

  async function combineFiles(files){
    var chunks = [];
    for(var i = 0; i < files.length; i += 1){
      var file = files[i];
      var text = await utils.readFileAsText(file);
      if(/\.csv$/i.test(file.name)){
        chunks.push(extractLinesFromCsv(text));
      } else {
        chunks.push(text);
      }
    }
    return chunks.filter(Boolean).join('\n');
  }

  function summarizeActiveState(events){
    var activeDigitsByName = new Map();
    var leftFinalDigitsByName = new Map();
    var nameOnlyJoin = new Set();

    function ensure(map, key){
      if(!map.has(key)) map.set(key, new Set());
      return map.get(key);
    }

    events.forEach(function(ev){
      var split = splitNickname(ev.nick);
      var activeSet;
      var leftSet;
      var marker;
      if(!split) return;
      marker = split.digits || NAME_ONLY_SENTINEL;
      if(ev.type === 'join'){
        ensure(activeDigitsByName, split.name).add(marker);
        if(!split.digits) nameOnlyJoin.add(split.name);
        if(leftFinalDigitsByName.has(split.name)){
          leftSet = leftFinalDigitsByName.get(split.name);
          leftSet.delete(marker);
          if(leftSet.size === 0) leftFinalDigitsByName.delete(split.name);
        }
      } else {
        activeSet = ensure(activeDigitsByName, split.name);
        activeSet.delete(marker);
        ensure(leftFinalDigitsByName, split.name).add(marker);
      }
    });

    return {
      activeDigitsByName: activeDigitsByName,
      leftFinalDigitsByName: leftFinalDigitsByName,
      nameOnlyJoin: nameOnlyJoin
    };
  }

  window.KakaoCheckParser = {
    NAME_ONLY_SENTINEL: NAME_ONLY_SENTINEL,
    splitNickname: splitNickname,
    parseChatText: parseChatText,
    combineFiles: combineFiles,
    summarizeActiveState: summarizeActiveState
  };
})();
