
/**
 * convert a number to letter
 */
function convertNum2Letter(num) {
  return String.fromCharCode(num + 65 - 1)
}

/**
 * convert a string(like 'TRUE'/'True'/'true') to boolean
 */
function convertBool(value) {
  return value.toLowerCase() == 'true';
}

/**
 * compare strings are the same, ignore upper or lower case
 * ex. 'Test' == 'TEST' is true
 */
function compareStrIgnoreCases(str1, str2) {
  flag = false;
  try {

    flag = str1.toLowerCase() == str2.toLowerCase();
  } catch (e) {
    Logger.log(e);
  }
  return flag;
}

/**
 * input raws:[][], schema:[]
 * return [jsonobject...]
 */
function convertRawsToJsons(raws, schema) {
  let jsons = [];
  raws.forEach(
    raw => {
      let json = {};
      // if (raw[17] == 'o') {
        // Logger.log(raw);
      // }
      for (let i = 0; i < schema.length; i++) {
        json[schema[i]] = raw[i];
      }
      jsons.push(json);
    }
  );
  return jsons;
}

/**
 * input json array, schema:[]
 * return [][]
 */
function convertJsonsToRaws(jsons, schema) {
  let raws = [];
  jsons.forEach(
    json => {
      let raw = [];
      for (let i = 0; i < schema.length; i++) {
        raw.push(json[schema[i]]);
      }
      raws.push(raw);
    }
  );
  return raws;
}

/**
 * input schema:[]
 * return an empty json object
 */
function getFormatJson(schema) {
  let json = {}
  schema.forEach(
    key => {
      json[key] = '';
    }
  )
  // Logger.log(json);
  return json;
}

/**
 * if str1 includes strs
 */
function includes(str1, str2) {
  let flag = false;
  if (str1) {
    flag = str1.includes(str2);
  }
  return flag;
}

/**
 * different days between 2 Date object
 */
function diffDays(date1, date2) {
  let diffTimestamp = date2.getTime() - date1.getTime();
  let diffDays =
    Math.round(diffTimestamp / (1000 * 3600 * 24));

  return diffDays;
}

/**
 * work days between 2 date(Date object)
 */
function countWorkdays(startDate, endDate) {
  let count = 0;
  let date1 = new Date(startDate);
  let date2 = new Date(endDate);
  while (date1 < date2) {
    date1.setDate(date1.getDate() + 1);
    if (!isWeekend(date1)) {
      count++;
    }
  }
  return count;
}

/**
 * if value is null or ''
 */
function isNull(value) {
  return value == null || value == '';
}

/**
 * if today is weekend
 */
function isWeekendToday() {
  return isWeekend(new Date());
}

/**
 * if a date is weekend
 */
function isWeekend(date) {
  return (date.getDay() == 0 || date.getDay() == 6);
}

/**
 * convert a day string to int for Date calculate
 */
function convertDayStr2Int(dayStr) {
  let day = -1;
  if (isNull(dayStr)) {
    return day;
  }
  if (dayStr.toLowerCase().includes('sun')) {
    day = 0;
  } else if (dayStr.toLowerCase().includes('mon')) {
    day = 1;
  } else if (dayStr.toLowerCase().includes('tue')) {
    day = 2;
  } else if (dayStr.toLowerCase().includes('wed')) {
    day = 3;
  } else if (dayStr.toLowerCase().includes('thu')) {
    day = 4;
  } else if (dayStr.toLowerCase().includes('fri')) {
    day = 5;
  } else if (dayStr.toLowerCase().includes('sat')) {
    day = 6;
  }
  return day;
}

/**
 * convert Month num from string like 'Jan' or 'Jun' to 0 or 5
 */
function getMonthNum(monStr) {
  let mon = -1;
  if (isNull(monStr)) {
    return mon;
  }
  if (monStr.toLowerCase().includes('jan')) {
    mon = 0;
  } else if (monStr.toLowerCase().includes('feb')) {
    mon = 1;
  } else if (monStr.toLowerCase().includes('mar')) {
    mon = 2;
  } else if (monStr.toLowerCase().includes('apr')) {
    mon = 3;
  } else if (monStr.toLowerCase().includes('may')) {
    mon = 4;
  } else if (monStr.toLowerCase().includes('jun')) {
    mon = 5;
  } else if (monStr.toLowerCase().includes('jul')) {
    mon = 6;
  } else if (monStr.toLowerCase().includes('aug')) {
    mon = 7;
  } else if (monStr.toLowerCase().includes('sep')) {
    mon = 8;
  } else if (monStr.toLowerCase().includes('oct')) {
    mon = 9;
  } else if (monStr.toLowerCase().includes('nov')) {
    mon = 10;
  } else if (monStr.toLowerCase().includes('dec')) {
    mon = 11;
  }
  return mon;
}

