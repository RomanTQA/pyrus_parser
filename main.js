const requests = require("./requests.js");

async function main() {
  try {
    //главный код

    //объявление переменных
    let flags = {
      isAuthDataRead: 1,
      isAuthDataValid: 1,
      isExcelFileFound: 1,
      isExcelFileRead: 1,
      isLinksArrFormed: 1,
      isLinksArrValid: 1,
      isIDsArrFormed: 1,
    };

    let isReadyToParse = {
      status: 1,
    };

    let flags2 = {
      isExcelEditedFileCreated: 1,
      isParseComplete: 1,
      isParseDataWritten: 1,
    };

    let stats = {
      richLinksCounter: 0,
      linksCounter: 0,
      pyrusLinkCounter: 0,
      idsCounter: 0,
      apiReqCounter: 0,
      apiReqTaskCounter: 0,
      apiResCounter: 0,
      apiTask200Counter: 0,
      cellTextCounter: 0,
      parsedCommCounter: 0,
    };

    let idsArray;
    let authData;
    let authRespCheck;
    let fileName;
    let flagsChecked;

    // 1. читаем файл настроек (может быть остановка кода)
    authData = await requests.readOrCreateAuth(flags);

    //Это вроде работает.

    //2. проверяем, что настройки валидные- security_key, login, auth_token
    authRespCheck = await requests.checkAuthData(flags, authData, stats);
    //3. ищем excel-файл в корневике
    fileName = requests.search_TaskList_To_Parse(flags);

    //5.парсим столбец А, и делаем из него idsArray
    idsArray = await requests.processFile(fileName, flags, stats);

    // определяем наблюдателя за isReadyToParse
    Object.defineProperty(isReadyToParse, "status", {
      get: function () {
        return isReadyToParse._status;
      },
      set: function (value) {
        isReadyToParse._status = value;
        if (isReadyToParse._status === true) {
          requests.startParserPrompt(
            authData,
            isReadyToParse,
            flags,
            flags2,
            fileName,
            idsArray,
            stats,
          );
        } else if (isReadyToParse._status === false) {
          requests.reporter(
            authData,
            isReadyToParse,
            flags,
            flags2,
            fileName,
            idsArray,
            stats,
          );
          requests.callWatcher();
        }
      },
    });
    //6. проверка готовности по флагам.
    await requests.waitForFlagsChange(flags, isReadyToParse);

    //здесь буду смотреть, что произошло
    //прототип второго наблюдателя
    Object.defineProperty(flags2, "isParseDataWritten", {
      get: function () {
        return flags2._isParseDataWritten;
      },
      set: function (value) {
        flags2._status = value;
        if (isReadyToParse._isParseDataWritten !== 1) {
          requests.reporter(
            authData,
            isReadyToParse,
            flags,
            flags2,
            fileName,
            idsArray,
            stats,
          );
        }
      },
    });
  } catch (error) {
    console.error(error);
  }
}
async function start() {
  await main();
}
start();
