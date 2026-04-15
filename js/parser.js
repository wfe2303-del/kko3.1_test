(function(){
  var utils = window.KakaoCheckUtils;
  var NAME_ONLY_SENTINEL = '__NAME_ONLY__';
  var JOIN_PATTERN = /^(.*?)님이\s*(입장하셨습니다|들어왔습니다|입장했습니다|들어오셨습니다)[\s.!]*$/;
  var LEAVE_PATTERN = /^(.*?)님이\s*(퇴장하셨습니다|나갔습니다|퇴장했습니다|퇴장하셨습니다|퇴장했습니다|퇴장하셨습니다|퇴장하였습니다)[\s.!]*$/;
  var KICK_PATTERN = /^(.*?)님을\s*(내보냈습니다|강퇴했습니다|추방했습니다)[\s.!]*$/;

  function splitNickname(nickname){
    var raw = String(nickname || '').trim();
    var digitIndex;
    var namePart;
    var digitsAll;
    var digits;
    var normalizedName;

    if(!raw) return null;

    digitIndex = raw.search(/\d/);
    namePart = digitIndex >= 0 ? raw.slice(0, digitIndex) : raw;
    normalizedName = utils.normalizeName(namePart);
    digitsAll = raw.replace(/\D+/g, '');
    digits = digitsAll.length >= 4 ? digitsAll.slice(-4) : null;

    if(!normalizedName){
      return {
        name: raw,
        digits: digits
      };
    }

    return {
      name: normalizedName,
      digits: digits
    };
  }

  function parseChatText(text){
    var lines = String(text || '').split(/\r?\n/);
    var events = [];
    var joinedCount = 0;
    var leftCount = 0;

    lines.forEach(function(rawLine){
      var line = String(rawLine || '').trim();
      var match;
      if(!line) return;

      match = line.match(JOIN_PATTERN);
      if(match){
        events.push({ type: 'join', nick: match[1].trim(), raw: line });
        joinedCount += 1;
        return;
      }

      match = line.match(LEAVE_PATTERN);
      if(match){
        events.push({ type: 'leave', nick: match[1].trim(), raw: line });
        leftCount += 1;
        return;
      }

      match = line.match(KICK_PATTERN);
      if(match){
        events.push({ type: 'leave', nick: match[1].trim(), raw: line });
        leftCount += 1;
      }
    });

    return {
      events: events,
      joinedCount: joinedCount,
      leftCount: leftCount
    };
  }

  function extractLinesFromCsv(text){
    var rows = utils.csvToRows(String(text || ''));
    var matched = [];
    var fallback = [];

    rows.forEach(function(row){
      if(!Array.isArray(row) || !row.length) return;

      row.forEach(function(cell, index){
        var value = String(cell || '').trim();
        if(!value) return;

        if(JOIN_PATTERN.test(value) || LEAVE_PATTERN.test(value) || KICK_PATTERN.test(value)){
          matched.push(value);
        }

        if(index === 2){
          fallback.push(value);
        }
      });
    });

    return matched.length ? matched.join('\n') : fallback.join('\n');
  }

  async function combineFiles(files){
    var chunks = [];
    var index;
    for(index = 0; index < files.length; index += 1){
      var file = files[index];
      var text = await utils.readFileAsText(file);
      chunks.push(/\.csv$/i.test(file.name) ? extractLinesFromCsv(text) : text);
    }
    return chunks.filter(Boolean).join('\n');
  }

  function summarizeActiveState(events){
    var activeDigitsByName = new Map();
    var leftFinalDigitsByName = new Map();

    function ensure(map, key){
      if(!map.has(key)) map.set(key, new Set());
      return map.get(key);
    }

    events.forEach(function(event){
      var split = splitNickname(event.nick);
      var marker;
      var activeSet;
      var leftSet;
      if(!split) return;

      marker = split.digits || NAME_ONLY_SENTINEL;

      if(event.type === 'join'){
        ensure(activeDigitsByName, split.name).add(marker);
        if(leftFinalDigitsByName.has(split.name)){
          leftSet = leftFinalDigitsByName.get(split.name);
          leftSet.delete(marker);
          if(leftSet.size === 0){
            leftFinalDigitsByName.delete(split.name);
          }
        }
        return;
      }

      activeSet = ensure(activeDigitsByName, split.name);
      activeSet.delete(marker);
      ensure(leftFinalDigitsByName, split.name).add(marker);
    });

    return {
      activeDigitsByName: activeDigitsByName,
      leftFinalDigitsByName: leftFinalDigitsByName
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
