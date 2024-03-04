const fs = require("fs");
const axios = require("axios"); //внешняя библиотека
const path = require("path");
const ExcelJS = require("exceljs"); //внешняя библиотека

//экспорт-модуль
module.exports = {
  readOrCreateAuth,
  get_Pyrus_Auth_Data,
  checkAuthData,
  check_Access_Token,
  search_TaskList_To_Parse,
  processFile,
  waitForFlagsChange,
  startParserPrompt,
  callWatcher,
  reporter,
  check_Access_Token,
  get_Access_Token,
  check_Links_Array,
  make_Ids_Array_To_Parse
};

//метод, который прочитает файл настроек, а если не получится- создаст дефолтный файл настроек и остановит код.
function readOrCreateAuth(flags) {
  const authData = get_Pyrus_Auth_Data(flags);
  if (authData === false) {
    const defaultAuthData = {
      security_key: "<YOUR_SECURITY_KEY>",
      login: "<YOUR_EMAIL_LOGIN>",
      access_token: "<YOUR_ACCESS_TOKEN>",
      fraud_delay: 700,
      matrix_mode: false,
      green_text: false,
	  extended_parse : false
    };
    const jsonData = JSON.stringify(defaultAuthData, null, 2);
    fs.writeFileSync("pyrus_auth.json", jsonData, "utf8");
    console.log(
      "Created new settings file 'pyrus_auth.json'. Please manually fill in your data - security_key, login for Pyrus.",
    );
    console.log(
      "Further execution of the code is meaningless without valid filled data in this file.",
    );
    process.exit(1);
  } else {
    console.log("User: ", authData.login);
    console.log("Fraud delay set to: ", authData.fraud_delay);
    if (authData.matrix_mode === true) {
      console.log("Matrix mode is active. Go fullscreen (recommended).");
    }
    if (authData.green_text === true) {
      console.log("Green text mode is active. Go fullscreen (recommended).");
    }
	if (authData.extended_parse ===true) {
		console.log("Extended parse option selected");
	}
    return authData;
  }
}
//++протестировал

//отдельный метод для чтения файла настроек pyrus_auth.json. устанавливает флаг isAuthDataRead- должен быть обьявлен зараннее.
function get_Pyrus_Auth_Data(flags) {
  try {
    const data = fs.readFileSync("pyrus_auth.json", "utf8");
    const jsonData = JSON.parse(data);
    flags.isAuthDataRead = true;
    return jsonData;
  } catch (err) {
    console.error("Error reading pyrus_auth.json file:", err);
    flags.isAuthDataRead = false;
    return false;
  }
}
//++ протестировал

//метод проверит токен, а если он плохой - отправит запрос на обновление токена. устанавливает флаг isAuthDataValid, он должен быть обьявлен зараннее
async function checkAuthData(flags, authData, stats) {
  let response = await check_Access_Token(authData, stats);

  if (response === 200) {
    flags.isAuthDataValid = true;
    return response;
  } else {
    flags.isAuthDataValid = false;
    console.error(`Error: ${response.error}`);

    if (response !== 200) {
      console.log("Error occurred. Trying to refresh access token...");
      let refreshResponse = await get_Access_Token(authData, stats);

      if (refreshResponse === 200) {
        console.log("Access token refreshed successfully.");
      } else {
        console.error(
          "Got error - check your Security_key or Login in settings file.",
        );
      }
    }

    return response;
  }
}
//++протестировал

//метод проверки, не протух ли текущий токен из настроек (api)
async function check_Access_Token(authData, stats) {
  stats.apiReqCounter++;
  try {
    const response = await axios.get(
      "https://api.pyrus.com/v4/profile?include_inactive=true",
      {
        headers: {
          Authorization: `Bearer ${authData.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );
    if (response.status === 200) {
      stats.apiResCounter++;
      if (response.data.first_name && response.data.last_name) {
        console.log(
          `Hi, ${response.data.first_name} ${response.data.last_name}!`,
        );
      }
    }

    return response.status;
  } catch (error) {
    console.error(`Error: ${error.message}`);
    return error.response.status;
  }
}
//++протестировал

//метод для получения access_token (api)
async function get_Access_Token(authData, stats) {
  stats.apiReqCounter++;

  // Отправка запроса на авторизацию
  return axios
    .post(
      "https://api.pyrus.com/v4/auth/",
      {
        login: authData.login,
        security_key: authData.security_key,
      },
      {
        headers: {
          "Content-Type": "application/json",
        },
      },
    )
    .then((response) => {
      const data = response.data;
      if (data.access_token) {
        stats.apiResCounter++;
        // Сохранение access_token в переменную и в файле pyrus_auth.json
        authData.access_token = data.access_token;
        fs.writeFileSync(
          "./pyrus_auth.json",
          JSON.stringify(authData, null, 2),
        );
        console.log("New access_token successfully saved");
        return 200;
      } else {
        console.error(
          `Error: No access_token in the response. Response code: ${response.status}`,
        );
        return 400;
      }
    })
    .catch((error) => {
      console.error(`Error: ${error.message}`);
      return 400;
    });
}
//++вроде протестировал

//метод для поиска импортного экселевского файла от паруса. Устанавливает флаг isExcelFileFound. Флаг должен быть объявлен зараннее.
function search_TaskList_To_Parse(flags) {
  const files = fs
    .readdirSync("./")
    .filter((file) => file.startsWith("TaskList_") && file.endsWith(".xlsx"));

  if (files.length > 0) {
    const fileName = files.sort()[0]; // Берем первый файл по алфавитному порядку
    console.log(`Found file to parse: ${fileName}`);
    flags.isExcelFileFound = true;
    return fileName;
  } else {
    console.log("File to parse not found");
    flags.isExcelFileFound = false;
    return false;
  }
}
//++вроде протестировал

//метод процессинга от именифайла к массиву id
async function processFile(fileName, flags, stats) {
  try {
    const richArray = await parse_Links_From_File(fileName, flags, stats);
    const linkArray = extractLinksFromRichArray(richArray, flags, stats);
    const isLinksArrValid = check_Links_Array(linkArray, flags, stats);
    if (isLinksArrValid) {
      const idsArray = make_Ids_Array_To_Parse(linkArray, flags, stats);
      return idsArray;
    } else {
      throw new Error("Invalid links array");
    }
  } catch (error) {
    console.error(error);
    return false;
  }
}
//протестировал

//метод для получения rich-text массива из файла импорта. устанавливает флаг isExcelFileRead. нужно обьявить его зараннее
async function parse_Links_From_File(fileName, flags, stats) {
  const richArray = [];
  const workbook = new ExcelJS.Workbook();

  // Определение функции extractLink()
  function extractLink(cellValue) {
    return cellValue;
  }

  try {
    await workbook.xlsx.readFile(fileName);
    const worksheet = workbook.getWorksheet(1);

    for (let i = 1; i <= worksheet.rowCount; i++) {
      const cellValue = worksheet.getCell(`A${i}`).value;
      if (cellValue) {
        stats.richLinksCounter++;
        const link = extractLink(cellValue);
        if (link) {
          richArray.push(link);
        } else {
          console.log(`Failed to parse link from A${i}`);
          richArray.push(false);
        }
      } else {
        continue; // Пропускаем пустую ячейку
      }
    }

    if (richArray.length === 0) {
      console.log("Nothing parsed, returned False");
      flags.isExcelFileRead = false;
      return false;
    } else {
      console.log("Links Array successfully created");
      flags.isExcelFileRead = true;
      return richArray;
    }
  } catch (error) {
    console.log("Error reading the file:", error);
    flags.isExcelFileRead = false;
    return false;
  }
}
//вроде протестировал

// метод выдергивания ссылок hyperlink из rich-text массива от парсера xlsx файла. Устанавливает флаг isLinksArrFormed

function extractLinksFromRichArray(richArray, flags, stats) {
  let linkArray = [];

  try {
    for (let item of richArray) {
      if (typeof item === "object" && item.hyperlink) {
        stats.linksCounter++;
        linkArray.push(item.hyperlink);
      } else {
        linkArray.push(false);
      }
    }

    flags.isLinksArrFormed = true; // Устанавливаем флаг в true при успешном выполнении метода
  } catch (error) {
    console.error("Ошибка при извлечении ссылок из rich-массива:", error);
    flags.isLinksArrFormed = false; // Устанавливаем флаг в false в случае ошибки
  }

  return linkArray;
}
//++ протестировал

//метод для анализа массива со ссылками ( дебаг) . ставит флаг isLinksArrValid (пришлось дублировать)
function check_Links_Array(linkArray, flags, stats) {
  if (linkArray.length === 0) {
    console.log("Links array is empty. Better to stop script...");
    flags.isLinksArrValid = false;
    return false;
  }

  let containsPyrusLink = false;
  let countFalse = 0;

  linkArray.forEach((element) => {
    if (typeof element === "string" && element.includes("https://pyrus.com/")) {
      stats.pyrusLinkCounter++;
      containsPyrusLink = true;
    } else if (element === false) {
      countFalse++;
    }
  });

  if (!containsPyrusLink) {
    console.log(
      "Links array has no links. Nothing to parse. Better to stop script...",
    );
    flags.isLinksArrValid = false;
    return false;
  }

  if (countFalse > 1) {
    console.log("Array has more than one boolean falses. Just for info.");
  }

  console.log("Links array contains links. Can try to parse them");
  flags.isLinksArrValid = true;
  return true;
}
// ++ протестировал

//вытащить id  задачи из links array
function make_Ids_Array_To_Parse(linkArray, flags, stats) {
  const idsArray = [];

  for (let i = 0; i < linkArray.length; i++) {
    const link = linkArray[i];
    if (link === false) {
      idsArray.push(false);
      console.log(
        "On iteration " + i + ", 'false' received and 'false' was written",
      );
    } else {
      const idMatch = link.match(/id(\d+)/);
      if (idMatch) {
        stats.idsCounter++;
        idsArray.push(parseInt(idMatch[1]));
      } else {
        idsArray.push(false);
        console.log("On iteration " + i + ", no id found, 'false' was written");
      }
    }
  }

  // Проверка на соответствие длины входного и выходного массивов
  if (idsArray.length !== linkArray.length) {
    flags.isIDsArrFormed = false;
    console.log("Input and output arrays have different lengths");
  } else {
    flags.isIDsArrFormed = true;
    console.log("id-s array formed");
  }

  return idsArray;
}

// ++ протестировал

//монитор флагов
function checkFlags(flags) {
  const flagValues = Object.values(flags);
  if (flagValues.includes(1)) {
    return false;
  }

  return flagValues;
}
//метод делает "десятисекундную" готовность парсить, путем установки флага isReadyToParse.status
function waitForFlagsChange(flags, isReadyToParse) {
  const interval = setInterval(() => {
    const result = checkFlags(flags);
    if (result.every((flag) => flag === true)) {
      clearInterval(interval);
      isReadyToParse.status = true;
      console.log("isReadyToParse status: true");
      clearTimeout(timeID);
    } else {
      if (result.some((flag) => flag === false)) {
        clearInterval(interval);
        isReadyToParse.status = false;
        console.log("isReadyToParse status: false");
        clearTimeout(timeID);
      }
    }
  }, 100);
  const timeID = setTimeout(() => {
    clearInterval(interval);
    isReadyToParse.status = false;
    console.log("isReadyToParse status: false, some flags are ===1");
  }, 10000); // Таймер на 10 секунд
}
//вроде бы Да

//разрешение парсить для вотчера Object.defineProperty
//разрешение парсить
async function startParserPrompt(
  authData,
  isReadyToParse,
  flags,
  flags2,
  fileName,
  idsArray,
  stats,
) {
  console.log("Ready to parse. Parse?  'y' or 'n'");

  process.stdin.setRawMode(true);
  process.stdin.resume();
  process.stdin.setEncoding("utf8");

  const handleInput = (data) => {
    if (data === "y" || data === "Y" || data === "\n") {
      copyAndRenameFile(fileName, flags2, (err, editedFileName) => {
        if (err) {
          console.error("Ошибка при копировании файла:", err);
        } else {
          console.log("Starting parser...");
		  
		  if (authData.extended_parse === true){
			  processParseExtended(idsArray, editedFileName, flags2, authData, stats);
		  } else {
          processParse(idsArray, editedFileName, flags2, authData, stats);
		  }
        }
      });

      process.stdin.removeListener("data", handleInput);
      process.stdin.pause();
    } else if (data === "n" || data === "N" || data === "\u001B") {
      reporter(
        authData,
        isReadyToParse,
        flags,
        flags2,
        fileName,
        idsArray,
        stats,
      );
      console.log("Stopping script...");
      process.exit(0);
    } else {
      console.log(
        "Invalid input. Please press 'y' to start Parse, or 'n' to stop script.",
      );
    }
  };

  process.stdin.on("data", handleInput);
}
//вроде бы Да

//метод создает копию файла xslx с приставкой _edited в конце. устанавливает флаг isExcelEditedFileCreated
function copyAndRenameFile(fileName, flags2, callback) {
  const filePath = `${__dirname}/${fileName}`;
  const editedFileName = `${fileName.replace(".xlsx", "_edited.xlsx")}`;
  const editedFilePath = `${__dirname}/${editedFileName}`;

  fs.copyFile(filePath, editedFilePath, (err) => {
    if (err) {
      console.error("Ошибка при копировании файла:", err);
      flags2.isExcelEditedFileCreated = false;
      callback(err, null);
    } else {
      console.log(`Файл успешно скопирован и переименован в ${editedFileName}`);
      flags2.isExcelEditedFileCreated = true;
      callback(null, editedFileName);
    }
  });
}
//++протестил

//процессинг: парсим и записываем в файл
async function processParse(idsArray, editedFileName, flags2, authData, stats) {
  try {
    // Шаг 1: Вызываем первый асинхронный метод parseTasks(idsArray)
    const parsedArray = await parseTasks(idsArray, flags2, authData, stats);

    // Шаг 2: Вызываем второй асинхронный метод writeDataToWorkbook(editedFileName, parsedArray)
    await writeDataToWorkbook(editedFileName, parsedArray, flags2, stats);

    console.log("Процесс обработки данных успешно завершен.");
  } catch (error) {
    console.error("Произошла ошибка в процессе обработки данных:", error);
  }
}

// Парсер
async function parseTasks(idsArray, flags2, authData, stats) {
  const parsedArray = [];

  try {
    for (let i = 0; i < idsArray.length; i++) {
      if (typeof idsArray[i] === "number") {
        const response = await get_Task(idsArray[i], authData, stats);

        if (response === false) {
          parsedArray.push(false);
          console.log(
            `On iteration ${i}, 'false' received, request skipped, 'false' was written`,
          );
        } else if (
          response.task &&
          response.task.comments &&
          response.task.comments[0] &&
          response.task.comments[0].text
        ) {
          stats.parsedCommCounter++;
          parsedArray.push(response.task.comments[0].text);
          if (authData.green_text === true) {
            console.log(response.task.comments[0].text);
          }
        } else {
          parsedArray.push("not found");
          console.log(
            `On iteration ${i}, comment was not found in the response body.`,
          );
        }

        // Добавляем паузу в 700 мс между запросами
        await new Promise((resolve) =>
          setTimeout(resolve, authData.fraud_delay),
        );
      } else {
        parsedArray.push(false);
        console.log(
          `On iteration ${i}, 'false' received, request skipped, 'false' was written`,
        );
      }
    }

    console.log("Parse array formed.");
    flags2.isParseComplete = true;
    return parsedArray;
  } catch (error) {
    console.error(`Error occurred during parsing: ${error.message}`);
    flags2.isParseComplete = false;
    return false;
  }
}
//ок

//метод , получающий всю задачу по ее id (api)
async function get_Task(taskId, authData, stats) {
  try {
    stats.apiReqCounter++;
    stats.apiReqTaskCounter++;
    const response = await axios.get(
      `https://api.pyrus.com/v4/tasks/${taskId}`,
      {
        headers: {
          Authorization: `Bearer ${authData.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );
    if (authData.matrix_mode === true) {
      console.log(JSON.stringify(response.data, null, 2));
    }
    if (response.status === 200) {
      stats.apiResCounter++;
      stats.apiTask200Counter++;
	  process.stdout.write(`\rAPI Task 200 Counter: ${stats.apiTask200Counter}`);
    }
    return response.data;
  } catch (error) {
    console.error(`Error: ${error.message}`);
    return error.response.status;
  }
}
// ++ протестировал.

//метод записывает  массив в файл в столбец I или в следующий доступный. устанавливает флаг isParseDataWritten
function writeDataToWorkbook(editedFileName, parsedArray, flags2, stats) {
  const workbook = new ExcelJS.Workbook();

  workbook.xlsx
    .readFile(editedFileName)
    .then(function () {
      const worksheet = workbook.getWorksheet("Список задач"); // Получаем лист по названию

      let columnIndex = 9; // Столбец I

      // Находим первый свободный столбец

      let column = worksheet.getColumn(columnIndex);
      let allCellsEmpty = column.values.every((value) => !value);

      while (!allCellsEmpty) {
        columnIndex++;
        column = worksheet.getColumn(columnIndex);
        allCellsEmpty = column.values.every((value) => !value);
      }
      worksheet.getColumn(columnIndex).width = 80; //установил ширину столбца 80 символов

      for (let i = 0; i < parsedArray.length; i++) {
        let cellValue = parsedArray[i];
        stats.cellTextCounter++;
        if (i === 0 && cellValue === false) {
          worksheet.getCell(1, columnIndex).value = "Comments: ";
        } else {
          worksheet.getCell(i + 1, columnIndex).value = cellValue;
          const lines = cellValue.split("\n");
          const rowCount = lines.length; //считаем высоту строки, по количеству строчек в тексте ячейки.
          worksheet.getRow(i + 1).height = rowCount * 15;
          worksheet.getCell(i + 1, columnIndex).alignment = {
            vertical: "top", // Выравнивание по верхнему краю
            wrapText: true, //перенос по словам
          };
        }
      }

      return workbook.xlsx.writeFile(editedFileName);
    })
    .then(function () {
      console.log("Данные успешно записаны в файл.");
      flags2.isParseDataWritten = true;
    })
    .catch(function (error) {
      console.error("Произошла ошибка при записи данных в файл:", error);
      flags2.isParseDataWritten = false;
    });
}

//++ протестировал

//инпут и закрытие для вотчера Object.defineProperty
function callWatcher() {
  console.log(
    "Parse data was formed with errors - cannot continue. Press any key. Stopping script.",
  );
  process.stdin.setRawMode(true);
  process.stdin.resume();
  process.stdin.on("data", process.exit.bind(process, 0));
}

//эмулятор парсера
function parser() {
  console.log(
    "parse...\n Parse...\n pArse...\n paRse...\n parSe...\n parsE...\n",
  );
}

//репортер
function reporter(
  authData,
  isReadyToParse,
  flags,
  flags2,
  fileName,
  idsArray,
  stats,
) {
  console.log("**********");
  console.log("Reporter Summary: ");
  console.log("");
  console.log("user: ", authData.login);
  console.log("ids Task list from: ", fileName);
  console.log("Approx. ids quantity: ", idsArray.length);
  console.log("Initialization flags status: ", flags);
  console.log("Was ready to parse: ", isReadyToParse);
  console.log("Parser flags status: ", flags2);
  console.log("Statistics counters: ", stats);
  console.log("**********");
}

// добавления

//вернет максимум длины массива-элемента
function getMaxLength(parsedArray){
	let max =0;
	for(let i=0; i < parsedArray.length; i++){
		if (parsedArray[i].length > max){
			max = parsedArray[i].length;
		}
		
	}
	return max;
}



// Парсер
async function parseTasksExtended(idsArray, flags2, authData, stats) {
  const parsedArray = [];

  try {
    for (let i = 0; i < idsArray.length; i++) {
      if (typeof idsArray[i] === "number") {
        const response = await get_Task(idsArray[i], authData, stats);

        if (response === false) {
          parsedArray.push(false);
          console.log(
            `On iteration ${i}, 'false' received, request skipped, 'false' was written`,
          );
        } else if (
          response.task &&
          response.task.comments 
        ) {
          stats.parsedCommCounter++;
          parsedArray.push(response.task.comments);
          if (authData.green_text === true) {
            console.log(response.task.comments);
          }
        } else {
          parsedArray.push("not found");
          console.log(
            `On iteration ${i}, comment was not found in the response body.`,
          );
        }

        // Добавляем паузу 
        await new Promise((resolve) =>
          setTimeout(resolve, authData.fraud_delay),
        );
      } else {
        parsedArray.push(false);
        console.log(
          `On iteration ${i}, 'false' received, request skipped, 'false' was written`,
        );
      }
    }

    console.log("Parse array formed.");
    flags2.isParseComplete = true;
  
    return parsedArray;
  } catch (error) {
    console.error(`Error occurred during parsing: ${error.message}`);
    flags2.isParseComplete = false;
    return false;
  }
}

//приводит данные в удобный вид для записи в файл

function transformTasksArray(tasksArray) {
  const result = [];

  for (const task of tasksArray) {
    if (!Array.isArray(task)) {
      result.push(false);
      continue;
    }

    const transformedTask = [];

    for (const event of task) {
      let transformedString = '';

      if (event.create_date) {
        transformedString +=` Дата: ${event.create_date}\n`;
      } else {
        transformedString += 'false\n';
      }

      if (event.author && event.author.first_name) {
        transformedString +=` by: ${event.author.first_name}`;
      } else {
        transformedString += 'false';
      }

      if (event.author && event.author.last_name) {
        transformedString +=`  ${event.author.last_name}\n`;
      } else {
        transformedString += '\n';
      }

      if (event.text) {
        transformedString +=` ${event.text}\n`;
      } else {
        transformedString += 'false\n';
      }

      if (event.reassigned_to) {
        transformedString += 'назначен на:\n';
      } else {
        transformedString += 'false\n';
      }

      if (event.reassigned_to && event.reassigned_to.first_name) {
        transformedString +=` ${event.reassigned_to.first_name}`;
      } else {
        transformedString += 'false';
      }

      if (event.reassigned_to && event.reassigned_to.last_name) {
        transformedString +=`  ${event.reassigned_to.last_name}`;
      } else {
        transformedString += '';
      }

      transformedTask.push(transformedString);
    }

    result.push(transformedTask);
  }

  return result;
}



//записывает "в строчку" массив

//метод записывает  массив в файл в столбец I или в следующий доступный. устанавливает флаг isParseDataWritten
function writeDataToWorkbookExtended(editedFileName, parsedArray, flags2, stats) {
  const workbook = new ExcelJS.Workbook();

  workbook.xlsx
    .readFile(editedFileName)
    .then(function () {
      const worksheet = workbook.getWorksheet("Список задач"); // Получаем лист по названию
	//узнать максимальную длину массива-элемента внутри parsedArray
	  let maxElemLength = getMaxLength(parsedArray);	
		
      let columnIndex = 9; // Столбец I
	  
      // Находим первый свободный столбец

      let column = worksheet.getColumn(columnIndex);
      let allCellsEmpty = column.values.every((value) => !value);

      while (!allCellsEmpty) {
        columnIndex++;
        column = worksheet.getColumn(columnIndex);
        allCellsEmpty = column.values.every((value) => !value);
      }
      worksheet.getColumn(columnIndex).width = 80; //установил ширину столбца 80 символов
		for (let i =1; i< maxElemLength; i++){
			worksheet.getColumn(columnIndex + i).width = 50;	//для остальных столбцов задействованных- ширина 50
		}

for (let i = 0; i < parsedArray.length; i++) {
    let cellValue = parsedArray[i];
    stats.cellTextCounter++;

    if (i === 0 && cellValue === false) {
        worksheet.getCell(1, columnIndex).value = "Comments: ";
    }

    if (Array.isArray(cellValue)) {
        let row = i + 1;
        if (cellValue.length > 0) {
            for (let j = 0; j < cellValue.length; j++) {
                worksheet.getCell(row, columnIndex + j).value = cellValue[j];
                worksheet.getCell(row, columnIndex + j).alignment = {
                    vertical: "top",
                    wrapText: true
                };
            }
        }
        const lines = cellValue[0].split("\n");
        const rowCount = lines.length;
        worksheet.getRow(row).height = rowCount * 15;
    } else {

        worksheet.getCell(i + 1, columnIndex).value = cellValue;
		
		if (i === 0 && cellValue === false) {
        worksheet.getCell(1, columnIndex).value = "Comments: ";
    }
        /* const lines = cellValue.split("\n");
        const rowCount = lines.length;
        worksheet.getRow(i + 1).height = rowCount * 15;
        worksheet.getCell(i + 1, columnIndex).alignment = {
            vertical: "top",
            wrapText: true 
        };*/
    }
}

      return workbook.xlsx.writeFile(editedFileName);
    })
    .then(function () {
      console.log("Данные успешно записаны в файл.");
      flags2.isParseDataWritten = true;
    })
    .catch(function (error) {
      console.error("Произошла ошибка при записи данных в файл:", error);
      flags2.isParseDataWritten = false;
    });
}


//процессинг: парсим и записываем в файл
async function processParseExtended(idsArray, editedFileName, flags2, authData, stats) {
  try {
    // Шаг 1: Вызываем первый асинхронный метод parseTasks(idsArray)
    const rawParsedArray = await parseTasksExtended(idsArray, flags2, authData, stats);
	
	//Шаг 1-а: Приводим массив к удобному виду
	const parsedArray = transformTasksArray(rawParsedArray);

    // Шаг 2: Вызываем второй асинхронный метод writeDataToWorkbook(editedFileName, parsedArray)
    await writeDataToWorkbookExtended(editedFileName, parsedArray, flags2, stats);

    console.log("Процесс обработки данных успешно завершен.");
  } catch (error) {
    console.error("Произошла ошибка в процессе обработки данных:", error);
  }
}