// On open function
function onOpen() {
  SpreadsheetApp.getActive().addMenu("KPIs", [
    { name: "Obtener datos", functionName: "trelloKPIs" },
  ]);
}

// Main function
function trelloKPIs() {
  // Constants
  const TRELLO_URL = "https://api.trello.com/1/";
  const TEST_LT = "test";
  const DONE_LISTS = ["Done this week"];
  const MONTH_NAMES = [
    "ene",
    "feb",
    "mar",
    "abr",
    "may",
    "jun",
    "jul",
    "ago",
    "sep",
    "oct",
    "nov",
    "dic",
  ];

  const CREDENTIALS = {
    KEY: "", // PUT YOUR TRELLO KEY HERE!
    TOKEN: "", // PUT YOUR TRELLO TOKEN HERE!
  };

  const STATUS = {
    PLANIFICADO: "planificado",
    ADICIONAL: "adicional",
    CERRADO: "cerrado",
    PENDIENTE: "pendiente",
  };

  const BOARDS = {
    PROJECT_1: {
      ID: "", // PUT TRELLO BOARD ID HERE!
      NAME: "test-project-1",
      LT: TEST_LT,
    },
    PROJECT_2: {
      ID: "", // PUT TRELLO BOARD ID HERE!
      NAME: "test-project-2",
      LT: TEST_LT,
    },
  };

  // Step 1: Get active sheet
  const getActiveSheet = () =>
    SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // Step 2: Set week number
  const setWeekNumber = () => {
    const result = SpreadsheetApp.getUi()
      .prompt("Ingrese nr de semana")
      .getResponseText();

    if (isNaN(result)) return;
    return Number(result);
  };

  // Step 3: Get board values
  const getBoardValues = (board) => {
    const boardUrl = `${TRELLO_URL}boards/${board.ID}/`;
    const urlParams = `?key=${CREDENTIALS.KEY}&token=${CREDENTIALS.TOKEN}`;
    const listsUrl = boardUrl + "lists" + urlParams;
    const cardsUrl = boardUrl + "cards" + urlParams;
    const membersUrl = boardUrl + "members" + urlParams;

    const fetchJson = (url) => {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
      }).getContentText();

      return JSON.parse(response);
    };

    return {
      lt: board.LT,
      name: board.NAME,
      lists: fetchJson(listsUrl),
      cards: fetchJson(cardsUrl),
      members: fetchJson(membersUrl),
    };
  };

  // Step 4: Write sheet
  const writeSheet = (week, sheet, boards) => {
    let rowCounter = 1;

    // Write rows
    const writeRow = (board) => {
      for (const card of board.cards) {
        if (!card.start || !card.due) continue;

        const startDate = new Date(card.start);
        const dueDate = new Date(card.due);
        const startWeek = getWeekNumber(startDate);
        const dueWeek = getWeekNumber(dueDate);
        if (week != startWeek || week != dueWeek) continue;

        // Increment
        rowCounter++;

        // Set values
        const responsible = board.members.find(
          ({ id }) => id == card.idMembers
        )?.fullName;
        const listName = board.lists
          .find(({ id }) => id == card.idList)
          .name.toUpperCase();
        const type = card.labels.find(({ name }) => name === STATUS.PLANIFICADO)
          ? STATUS.PLANIFICADO
          : STATUS.ADICIONAL;
        const state = DONE_LISTS.includes(listName)
          ? STATUS.CERRADO
          : STATUS.PENDIENTE;
        const hours = getHours(card.name);
        const hoursPla = type === STATUS.PLANIFICADO ? hours : 0;
        const hoursWor = state === STATUS.CERRADO ? hours : 0;
        const hoursDon =
          type === STATUS.PLANIFICADO && state === STATUS.CERRADO ? hours : 0;
        const hoursPen =
          type === STATUS.PLANIFICADO && state === STATUS.PENDIENTE ? hours : 0;
        const hoursAdd = type === STATUS.ADICIONAL ? hours : 0;

        // Write values on sheet
        sheet.getRange(rowCounter, 1).setValue(card.shortUrl);
        sheet.getRange(rowCounter, 2).setValue(card.name);
        sheet.getRange(rowCounter, 3).setValue(board.name);
        sheet.getRange(rowCounter, 4).setValue(responsible);
        sheet.getRange(rowCounter, 5).setValue(startDate.getFullYear());
        sheet.getRange(rowCounter, 6).setValue(startDate.getMonth() + 1);
        sheet
          .getRange(rowCounter, 7)
          .setValue(MONTH_NAMES[startDate.getMonth()]);
        sheet.getRange(rowCounter, 8).setValue(startWeek);
        sheet.getRange(rowCounter, 9).setValue(startDate);
        sheet.getRange(rowCounter, 10).setValue(dueDate);
        sheet.getRange(rowCounter, 11).setValue(type);
        sheet.getRange(rowCounter, 12).setValue(state);
        sheet.getRange(rowCounter, 13).setValue(hours);
        sheet.getRange(rowCounter, 14).setValue(hoursPla);
        sheet.getRange(rowCounter, 15).setValue(hoursWor);
        sheet.getRange(rowCounter, 16).setValue(hoursDon);
        sheet.getRange(rowCounter, 17).setValue(hoursPen);
        sheet.getRange(rowCounter, 18).setValue(hoursAdd);
        sheet.getRange(rowCounter, 19).setValue(board.lt);
      }
    };

    // Get week number
    const getWeekNumber = (date) => {
      const firstDayOfYear = new Date(new Date().getFullYear(), 0, 0); // Depending on the current year, switch last arg to 0 or 1.
      const firstOp = date - firstDayOfYear;
      const secondOp = firstOp / (24 * 60 * 60 * 1000);
      const thirdOp = Math.floor(secondOp);
      const fourthOp = thirdOp / 7;

      return Math.ceil(fourthOp);
    };

    // Get hours
    const getHours = (name) => {
      start = name.lastIndexOf("(");
      end = name.lastIndexOf(")");

      if (start >= 0 && end >= 0)
        return name.substring(start + 1, end).replace(",", ".");
      return;
    };

    // Repeat on each board
    boards.forEach((board) => writeRow(board));
  };

  // Step 5: Execute everything
  const sheet = getActiveSheet();
  const week = setWeekNumber();
  const boards = [
    getBoardValues(BOARDS.PROJECT_1),
    getBoardValues(BOARDS.PROJECT_2),
  ];

  writeSheet(week, sheet, boards);
}
