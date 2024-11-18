function newGame() {
  stopGame();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Opponent");
  const boardSize = 10;
  const ships = [4, 3, 3, 2, 2, 2, 1, 1, 1, 1];

  sheet.getRange('C6:L15').clearContent();
  let board = Array.from({ length: boardSize }, () => Array(boardSize).fill(0));

  const canPlaceShip = (row, col, size, isHorizontal) => {
    for (let i = 0; i < size; i++) {
      const r = row + (isHorizontal ? 0 : i);
      const c = col + (isHorizontal ? i : 0);

      if (r >= boardSize || c >= boardSize || board[r][c] !== 0) return false;

      for (let x = -1; x <= 1; x++) {
        for (let y = -1; y <= 1; y++) {
          const nr = r + x, nc = c + y;
          if (nr >= 0 && nr < boardSize && nc >= 0 && nc < boardSize && board[nr][nc] !== 0) {
            return false;
          }
        }
      }
    }
    return true;
  };

  const placeShip = size => {
    while (true) {
      const isHorizontal = Math.random() < 0.5;
      const row = Math.floor(Math.random() * (boardSize - (isHorizontal ? 0 : size)));
      const col = Math.floor(Math.random() * (boardSize - (isHorizontal ? size : 0)));

      if (canPlaceShip(row, col, size, isHorizontal)) {
        for (let i = 0; i < size; i++) {
          board[row + (isHorizontal ? 0 : i)][col + (isHorizontal ? i : 0)] = size;
        }
        break;
      }
    }
  };

  ships.forEach(placeShip);

  board.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (cell !== 0) sheet.getRange(r + 6, c + 3).setValue('⛵');
    });
  });

  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Game").getRange("C4");
  
  var richText = SpreadsheetApp.newRichTextValue()
    .setText("Status: In Progress")
    .setTextStyle(0, 7, SpreadsheetApp.newTextStyle().setBold(false).build())
    .setTextStyle(8, 19, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("blue").build())
    .build();
  
  cell.setRichTextValue(richText);
}

function stopGame() {
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Game").getRange("C4");
  
  var richText = SpreadsheetApp.newRichTextValue()
    .setText("Status: Not Started")
    .setTextStyle(0, 7, SpreadsheetApp.newTextStyle().setBold(false).build())
    .setTextStyle(8, 19, SpreadsheetApp.newTextStyle().setBold(true).build())
    .build();
  
  cell.setRichTextValue(richText);

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Game").getRange('C6:L15').clearContent();
}

function shoot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName("Game");
  const opponentSheet = ss.getSheetByName("Opponent");

  const gameRange = gameSheet.getActiveCell();
  const gameRow = gameRange.getRow();
  const gameCol = gameRange.getColumn();

  if (gameRow < 6 || gameRow > 15 || gameCol < 3 || gameCol > 12) {
    SpreadsheetApp.getUi().alert("Выберите ячейку в игровом поле!");
    return;
  }

  const opponentCellValue = opponentSheet.getRange(gameRow, gameCol).getValue();
  const gameCell = gameSheet.getRange(gameRow, gameCol);

  if (gameCell.getValue() === '💥' || gameCell.getValue() === '◾️') {
    SpreadsheetApp.getUi().alert("Вы уже стреляли в эту ячейку!");
    return;
  }

  if (opponentCellValue === '⛵') {
    gameCell.setValue('💥');
    if (markSunkShipIfNeeded(gameSheet, opponentSheet, gameRow, gameCol)) {
      SpreadsheetApp.getUi().alert("Корабль уничтожен!");
    }
  } else {
    gameCell.setValue('◾️');
  }

  if (areAllShipsSunk(opponentSheet, gameSheet)) {
    SpreadsheetApp.getUi().alert("Игра завершена! Все корабли уничтожены!");
  
    var richText = SpreadsheetApp.newRichTextValue()
      .setText("Status: Completed")
      .setTextStyle(0, 7, SpreadsheetApp.newTextStyle().setBold(false).build())
      .setTextStyle(8, 17, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("green").build())
      .build();
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Game").getRange("C4").setRichTextValue(richText);

  }
}

function markSunkShipIfNeeded(gameSheet, opponentSheet, row, col) {
  const boardSize = 10;
  const rangeStartRow = 6;
  const rangeStartCol = 3;

  const visited = {};
  let isSunk = true;

  const traverseShip = (r, c) => {
    if (
      r < rangeStartRow || r >= rangeStartRow + boardSize ||
      c < rangeStartCol || c >= rangeStartCol + boardSize ||
      visited[r + ',' + c] || opponentSheet.getRange(r, c).getValue() !== '⛵'
    ) {
      return;
    }

    visited[r + ',' + c] = true;

    const gameCell = gameSheet.getRange(r, c).getValue();
    if (gameCell !== '💥') {
      isSunk = false;
    }

    traverseShip(r - 1, c);
    traverseShip(r + 1, c);
    traverseShip(r, c - 1);
    traverseShip(r, c + 1);
  };

  traverseShip(row, col);

  if (isSunk) {
    Object.keys(visited).forEach(key => {
      const [r, c] = key.split(',').map(Number);
      gameSheet.getRange(r, c).setValue('✖️');
    });
  }

  return isSunk;
}

function areAllShipsSunk(opponentSheet, gameSheet) {
  const boardSize = 10;
  const rangeStartRow = 6;
  const rangeStartCol = 3;

  for (let row = rangeStartRow; row < rangeStartRow + boardSize; row++) {
    for (let col = rangeStartCol; col < rangeStartCol + boardSize; col++) {
      const opponentCell = opponentSheet.getRange(row, col).getValue();
      const gameCell = gameSheet.getRange(row, col).getValue();

      if (opponentCell === '⛵' && gameCell !== '💥' && gameCell !== '✖️') {
        return false;
      }
    }
  }

  return true;
}
