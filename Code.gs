const SORTED_COLOR_IDENTITY_NAMES = {
  "": "Colorless",
  "W": "White",
  "U": "Blue",
  "B": "Black",
  "R": "Red",
  "G": "Green",
  "UW": "Azorius",
  "UB": "Dimir",
  "BR": "Rakdos",
  "RG": "Gruul",
  "GW": "Selesnya",
  "BW": "Orzhov",
  "UR": "Izzet",
  "BG": "Golgari",
  "RW": "Boros",
  "GU": "Simic",
  "BUW": "Esper",
  "BUR": "Grixis",
  "BGR": "Jund",
  "GRW": "Naya",
  "BGW": "Abzan",
  "RUW": "Jeskai",
  "BGU": "Sultai",
  "GUW": "Bant",
  "BRW": "Mardu",
  "GRU": "Temur",
  "BRUW": "Glint",
  "BRGU": "Yore",
  "BGUW": "Witch",
  "GRUW": "Ink",
  "BGRW": "Dune",
  "BGURW": "Five-Color"
};

function fetchColorIdentity(commanderName) {
  if (!commanderName) return "Unknown";

  const partnerNames = commanderName.split("+").map(name => name.trim());
  const colors = [];

  for (const name of partnerNames) {
    const url = "https://api.scryfall.com/cards/named?fuzzy=" + encodeURIComponent(name);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) continue;

    const card = JSON.parse(response.getContentText());
    if (Array.isArray(card.color_identity)) {
      colors.push(...card.color_identity);
    }
  }

  const sortedKey = [...new Set(colors)].sort().join("");
  return SORTED_COLOR_IDENTITY_NAMES[sortedKey] || "Unknown";
}


/**
 * MTG Stats Tracker - Google Apps Script
 * 
 * This script analyzes gameplay data from multiple sheets to generate comprehensive
 * statistics including player performance, commander win rates, seat position stats,
 * and player rivalries.
 */

function updateStats() {
  // Get the active spreadsheet and all sheets within it
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Define the name of our stats sheet and template sheets to ignore
  const statsSheetName = "Stats";
  const templateSheetNames = ["Template", "Template Sheet", "Data Template"];

  // Get or create the stats sheet
  let statsSheet = ss.getSheetByName(statsSheetName);
  if (!statsSheet) {
    statsSheet = ss.insertSheet(statsSheetName);
  } else {
    statsSheet.clear(); // Clear existing content if sheet exists
  }

  // Define headers for each section of the report
  const playerHeaders = [
    "Player",
    "Games Played",
    "Wins",
    "Win Rate (%)",
    "Avg Finish",
    "Top 3 Finishes",
    "First Bloods Dealt",
    "First Bloods Taken",
    "Top 1 Commander",
    "Top 2 Commander",
    "Top 1 Target (FB)",
    "Top 2 Target (FB)"
  ];

  const commanderHeaders = [
    "Commander",
    "Games Played",
    "Wins",
    "Win Rate (%)"
  ];

  const seatPositionHeaders = [
    "Seat Position",
    "Games Played",
    "Wins",
    "Win Rate (%)",
    "Avg Finish"
  ];

  const rivalryHeaders = [
    "Player",
    "Top Victim",
    "Nemesis"
  ];

  // Start with player stats section
  statsSheet.appendRow(playerHeaders);

  // Initialize objects to store our statistics
  const playerStats = {};
  const commanderStats = {};
  const seatPositionStats = {};

  // Process each game sheet to collect data
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name === statsSheetName || templateSheetNames.includes(name)) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];

    // Find column indices for required data
    const playerCol = headers.indexOf("Player");
    const commanderCol = headers.indexOf("Commander");
    const finishCol = headers.indexOf("Finish");
    const winCol = headers.indexOf("Win");
    const fbDealtCol = headers.indexOf("First Blood Dealt");
    const fbTakenCol = headers.indexOf("First Blood Taken");
    const diedToCol = headers.indexOf("Died To");
    const seatPosCol = headers.indexOf("Seat Pos");

    if ([playerCol, commanderCol, finishCol, winCol, fbDealtCol, fbTakenCol, diedToCol].includes(-1)) {
      Logger.log(`Sheet "${name}" is missing required columns.`);
      continue;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const player = row[playerCol];
      const commander = row[commanderCol];
      const finish = Number(row[finishCol]);
      const win = (row[winCol] ? row[winCol].toString() : "").toUpperCase() === "Y";
      const fbDealt = row[fbDealtCol] ? row[fbDealtCol].toString() : "";
      const fbTaken = row[fbTakenCol] ? row[fbTakenCol].toString() : "";
      const diedTo = row[diedToCol] ? row[diedToCol].toString() : "";
      const seatPos = seatPosCol !== -1 ? row[seatPosCol] : null;

      if (!player || !commander) continue;

      if (!playerStats[player]) {
        playerStats[player] = {
          games: 0,
          wins: 0,
          finishes: [],
          top3: 0,
          firstBloodsDealt: 0,
          firstBloodsTaken: 0,
          commanderCounts: {},
          fbTargets: {},
          victimCounts: {},
          nemesisCounts: {}
        };
      }

      const stats = playerStats[player];
      stats.games++;
      stats.wins += win ? 1 : 0;
      stats.finishes.push(finish);
      if (finish <= 3) stats.top3++;

      stats.commanderCounts[commander] = (stats.commanderCounts[commander] || 0) + 1;

      if (fbDealt) {
        stats.firstBloodsDealt++;
        stats.fbTargets[fbDealt] = (stats.fbTargets[fbDealt] || 0) + 1;
      }

      if (fbTaken) {
        stats.firstBloodsTaken++;
      }

      if (diedTo) {
        if (!playerStats[diedTo]) {
          playerStats[diedTo] = {
            games: 0,
            wins: 0,
            finishes: [],
            top3: 0,
            firstBloodsDealt: 0,
            firstBloodsTaken: 0,
            commanderCounts: {},
            fbTargets: {},
            victimCounts: {},
            nemesisCounts: {}
          };
        }
        playerStats[diedTo].victimCounts[player] = (playerStats[diedTo].victimCounts[player] || 0) + 1;
        stats.nemesisCounts[diedTo] = (stats.nemesisCounts[diedTo] || 0) + 1;
      }

      if (!commanderStats[commander]) {
        commanderStats[commander] = { games: 0, wins: 0 };
      }
      commanderStats[commander].games++;
      commanderStats[commander].wins += win ? 1 : 0;
      
      if (seatPosCol !== -1 && seatPos) {
        const seatPosStr = seatPos.toString();
        if (!seatPositionStats[seatPosStr]) {
          seatPositionStats[seatPosStr] = {
            games: 0,
            wins: 0,
            finishes: []
          };
        }
        seatPositionStats[seatPosStr].games++;
        seatPositionStats[seatPosStr].wins += win ? 1 : 0;
        seatPositionStats[seatPosStr].finishes.push(finish);
      }
    }
  }

  // Write player statistics to the sheet
  for (const player in playerStats) {
    const stats = playerStats[player];
    const avgFinish = stats.finishes.reduce((a, b) => a + b, 0) / stats.finishes.length;
    const winRate = (stats.wins / stats.games) * 100;

    const topCommanders = Object.entries(stats.commanderCounts).sort((a, b) => b[1] - a[1]).map(entry => entry[0]);
    const top1 = topCommanders[0] || "";
    const top2 = topCommanders[1] || "";

    const topTargets = Object.entries(stats.fbTargets).sort((a, b) => b[1] - a[1]).map(entry => entry[0]);
    const topTarget1 = topTargets[0] || "";
    const topTarget2 = topTargets[1] || "";

    statsSheet.appendRow([
      player,
      stats.games,
      stats.wins,
      winRate.toFixed(1),
      avgFinish.toFixed(2),
      stats.top3,
      stats.firstBloodsDealt,
      stats.firstBloodsTaken,
      top1,
      top2,
      topTarget1,
      topTarget2
    ]);

      const playerHeaderRange = statsSheet.getRange(1, 1, 1, playerHeaders.length);
  playerHeaderRange.setFontWeight("bold").setBackground("#1a73e8").setFontColor("white").setHorizontalAlignment("center");

  if (Object.keys(playerStats).length > 0) {
    const playerDataRange = statsSheet.getRange(2, 1, Object.keys(playerStats).length, playerHeaders.length);
    playerDataRange.setBackgrounds(generateStripedRows(Object.keys(playerStats).length, "#f1f3f4", "#ffffff", playerHeaders.length));
    playerDataRange.setFontWeight("bold").setHorizontalAlignment("center");
  }
  }

  // Add spacing after player stats section
  const lastPlayerRow = statsSheet.getLastRow();
  statsSheet.insertRowAfter(lastPlayerRow);
  statsSheet.setRowHeight(lastPlayerRow + 1, 25);
  statsSheet.getRange(lastPlayerRow + 1, 1).setValue("\u00A0");

  // COMMANDER STATS SECTION (left side)
  const commanderTitleRow = statsSheet.getLastRow() + 1;
  statsSheet.getRange(commanderTitleRow, 1).setValue("👑 Commander Win Rates").setFontWeight("bold").setFontSize(12);
  statsSheet.appendRow(commanderHeaders);

  const sortedCommanders = Object.entries(commanderStats)
    .map(([commander, stats]) => ({
      commander,
      games: stats.games,
      wins: stats.wins,
      winRate: (stats.wins / stats.games) * 100
    }))
    .sort((a, b) => b.winRate - a.winRate);

  for (const entry of sortedCommanders) {
    statsSheet.appendRow([
      entry.commander,
      entry.games,
      entry.wins,
      entry.winRate.toFixed(1)
    ]);
  }

  const commanderHeaderRange = statsSheet.getRange(commanderTitleRow + 1, 1, 1, commanderHeaders.length);
  commanderHeaderRange.setFontWeight("bold").setBackground("#1a73e8").setFontColor("white").setHorizontalAlignment("center");

  const commanderDataRange = statsSheet.getRange(commanderTitleRow + 2, 1, sortedCommanders.length, commanderHeaders.length);
  commanderDataRange.setBackgrounds(generateStripedRows(sortedCommanders.length, "#f1f3f4", "#ffffff", commanderHeaders.length));
  commanderDataRange.setFontWeight("bold").setHorizontalAlignment("center");

  // SEAT POSITION STATS SECTION (right side of commander stats)
  const seatPositionStartCol = commanderHeaders.length + 2;
  let seatPositionEndRow = commanderTitleRow + 1; // Initialize end row
  
  if (Object.keys(seatPositionStats).length > 0) {
    statsSheet.getRange(commanderTitleRow, seatPositionStartCol)
      .setValue("💺 Seat Position Stats")
      .setFontWeight("bold")
      .setFontSize(12);
    
    const seatPositionHeaderRange = statsSheet.getRange(
      commanderTitleRow + 1, 
      seatPositionStartCol, 
      1, 
      seatPositionHeaders.length
    );
    seatPositionHeaderRange.setValues([seatPositionHeaders]);
    seatPositionHeaderRange.setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("white")
      .setHorizontalAlignment("center");

    const sortedSeatPositions = Object.entries(seatPositionStats)
      .map(([seatPos, stats]) => ({
        seatPos,
        games: stats.games,
        wins: stats.wins,
        winRate: (stats.wins / stats.games) * 100,
        avgFinish: stats.finishes.reduce((a, b) => a + b, 0) / stats.finishes.length
      }))
      .sort((a, b) => parseInt(a.seatPos) - parseInt(b.seatPos));

    const seatPositionData = sortedSeatPositions.map(entry => [
      entry.seatPos,
      entry.games,
      entry.wins,
      entry.winRate.toFixed(1),
      entry.avgFinish.toFixed(2)
    ]);
    
    if (seatPositionData.length > 0) {
      const seatPositionDataRange = statsSheet.getRange(
        commanderTitleRow + 2, 
        seatPositionStartCol, 
        seatPositionData.length, 
        seatPositionHeaders.length
      );
      seatPositionDataRange.setValues(seatPositionData);
      seatPositionDataRange.setBackgrounds(
        generateStripedRows(
          seatPositionData.length, 
          "#f1f3f4", 
          "#ffffff", 
          seatPositionHeaders.length
        )
      );
      seatPositionDataRange.setFontWeight("bold").setHorizontalAlignment("center");
      
      seatPositionEndRow = commanderTitleRow + 1 + seatPositionData.length;
    }
  }

// Calculate where Player Rivalries should start
const seatPositionDataLength = Object.keys(seatPositionStats).length > 0 ? 
    Object.keys(seatPositionStats).length : 0;
    
const rivalryStartRow = commanderTitleRow + 2 + seatPositionDataLength;

// Add a single small spacer row ONLY if we have seat position data
if (seatPositionDataLength > 0) {
    statsSheet.getRange(rivalryStartRow, seatPositionStartCol)
        .setValue("") // Empty spacer cell
        .setBackground("#f1f3f4"); // Match your stripe color
    statsSheet.setRowHeight(rivalryStartRow, 10); // Very small spacer
}

// Player Rivalries starts immediately after spacer (or at commander level if no seat data)
const rivalryTitleRow = seatPositionDataLength > 0 ? rivalryStartRow + 1 : commanderTitleRow;

// Add Player Rivalries title
statsSheet.getRange(rivalryTitleRow, seatPositionStartCol)
    .setValue("\uD83D\uDD25 Player Rivalries")
    .setFontWeight("bold")
    .setFontSize(12);

// Add headers on next row
const rivalryHeaderRange = statsSheet.getRange(
    rivalryTitleRow + 1,
    seatPositionStartCol,
    1,
    rivalryHeaders.length
);
rivalryHeaderRange.setValues([rivalryHeaders]);
rivalryHeaderRange.setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("white")
    .setHorizontalAlignment("center");

// Add rivalry data
const rivalryData = [];
for (const player in playerStats) {
    const stats = playerStats[player];
    const topVictim = Object.entries(stats.victimCounts).sort((a, b) => b[1] - a[1]).map(entry => entry[0])[0] || "";
    const nemesis = Object.entries(stats.nemesisCounts).sort((a, b) => b[1] - a[1]).map(entry => entry[0])[0] || "";
    rivalryData.push([player, topVictim, nemesis]);
}

if (rivalryData.length > 0) {
    const rivalryDataRange = statsSheet.getRange(
        rivalryTitleRow + 2,
        seatPositionStartCol,
        rivalryData.length,
        rivalryHeaders.length
    );
    rivalryDataRange.setValues(rivalryData);
    rivalryDataRange.setBackgrounds(
        generateStripedRows(
            rivalryData.length,
            "#f1f3f4",
            "#ffffff",
            rivalryHeaders.length
        )
    );
    rivalryDataRange.setFontWeight("bold")
        .setHorizontalAlignment("center");
}

// --- COLOR IDENTITY TABLES BELOW PLAYER RIVALRIES ---

// Prepare color identity data first
const identityCounts = {};
const playerColorData = [];

for (const player in playerStats) {
  const stats = playerStats[player];
  const topCommanders = Object.entries(stats.commanderCounts)
    .sort((a, b) => b[1] - a[1])
    .map(e => e[0]);
  const topCommander = topCommanders[0];
  let identity = "Unknown";

  try {
    identity = fetchColorIdentity(topCommander);
  } catch (e) {
    identity = "Not Found";
  }

  playerColorData.push([player, identity]);
  identityCounts[identity] = (identityCounts[identity] || 0) + 1;
}

// Find the row just after Player Rivalries data to start color identity tables
const colorTableStartRow = rivalryTitleRow + 2 + rivalryData.length + 1; // only 1 row spacing now

const summaryStartCol = seatPositionStartCol; // keep summary at same col
const popularityStartCol = summaryStartCol + 3; // move popularity 3 cols to right (2 cols data + 1 col spacing)

// Player favorite color identity table header
statsSheet.getRange(colorTableStartRow, summaryStartCol)
  .setValue("🎨 Color Identity Summary")
  .setFontWeight("bold")
  .setFontSize(12);

// Favorite color identity table headers
const colorHeaders = ["Player", "Favorite Color Identity"];
statsSheet.getRange(colorTableStartRow + 1, summaryStartCol, 1, colorHeaders.length).setValues([colorHeaders])
  .setFontWeight("bold")
  .setBackground("#1a73e8")
  .setFontColor("white")
  .setHorizontalAlignment("center");

// Write player favorite color identity data
statsSheet.getRange(colorTableStartRow + 2, summaryStartCol, playerColorData.length, 2).setValues(playerColorData);

// Format data rows
statsSheet.getRange(colorTableStartRow + 2, summaryStartCol, playerColorData.length, 2)
  .setBackgrounds(generateStripedRows(playerColorData.length, "#f1f3f4", "#ffffff", 2))
  .setFontWeight("bold")
  .setHorizontalAlignment("center");

// Totals breakdown header
statsSheet.getRange(colorTableStartRow, popularityStartCol)
  .setValue("🏆 Color Identity Popularity")
  .setFontWeight("bold")
  .setFontSize(12);

// Totals table headers
const totalHeaders = ["Color Identity", "Count", "Percentage"];
statsSheet.getRange(colorTableStartRow + 1, popularityStartCol, 1, totalHeaders.length)
  .setValues([totalHeaders])
  .setFontWeight("bold")
  .setBackground("#1a73e8")
  .setFontColor("white")
  .setHorizontalAlignment("center");

// Build totals summary rows
const totalPlayers = playerColorData.length;
const sortedIdentities = Object.entries(identityCounts).sort((a, b) => b[1] - a[1]);
const summaryRows = sortedIdentities.map(([id, count]) => [id, count, ((count / totalPlayers) * 100).toFixed(1) + "%"]);

// Write totals data
statsSheet.getRange(colorTableStartRow + 2, popularityStartCol, summaryRows.length, totalHeaders.length).setValues(summaryRows);

// Format totals data rows
statsSheet.getRange(colorTableStartRow + 2, popularityStartCol, summaryRows.length, totalHeaders.length)
  .setBackgrounds(generateStripedRows(summaryRows.length, "#f1f3f4", "#ffffff", totalHeaders.length))
  .setFontWeight("bold")
  .setHorizontalAlignment("center");

  row = statsSheet.getLastRow() + 2;


  // Final formatting and cleanup
  const columnWidths = {
    "Games Played": 95,
    "Wins": 80,
    "Player": 275,
    "Commander": 140,
    "Win Rate (%)": 100,
    "Avg Finish": 90,
    "First Bloods Dealt": 145,
    "First Bloods Taken": 130,
    "Top 1 Commander": 200,
    "Top 2 Commander": 200,
    "Top 1 Target (FB)": 160,
    "Top 2 Target (FB)": 160,
    "Top Victim": 160,
    "Nemesis": 160,
    "Seat Position": 100
  };

  const lastCol = statsSheet.getLastColumn();
  for (let i = 1; i <= lastCol; i++) {
    statsSheet.autoResizeColumn(i);
    const header = statsSheet.getRange(1, i).getValue();
    if (columnWidths[header]) {
      statsSheet.setColumnWidth(i, columnWidths[header]);
    }
  }

  statsSheet.setRowHeights(1, statsSheet.getMaxRows(), 25);

  // --- CURSOR PARKING LOT (with Blue Header and Padding) ---
  const parkingLotLabel = "AFK";
  const parkingLotTotalRows = 7; // 1 header + 6 blank rows
  const parkingLotTotalCols = 1;

  const parkingLotStartRow = rivalryTitleRow + 1; // move 1 row down from Rivalries title
  const parkingLotStartCol = seatPositionStartCol + rivalryHeaders.length + 1; // 1 column to the right of Rivalries

  // Header (blue styled)
  statsSheet.getRange(parkingLotStartRow, parkingLotStartCol)
    .setValue(parkingLotLabel)
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Determine how many data rows are in the Player Rivalries section
  const rivalryDataRowCount = rivalryData.length;

  // Generate same number of blank rows below the header
  const dynamicParkingLotRows = Array(rivalryDataRowCount).fill([""]);
  statsSheet.getRange(parkingLotStartRow + 1, parkingLotStartCol, dynamicParkingLotRows.length, parkingLotTotalCols)
    .setValues(dynamicParkingLotRows)
    .setBackgrounds(generateStripedRows(dynamicParkingLotRows.length, "#f1f3f4", "#ffffff", parkingLotTotalCols))
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Box around the header + dynamic rows
  const boxRange = statsSheet.getRange(parkingLotStartRow, parkingLotStartCol, 1 + rivalryDataRowCount, parkingLotTotalCols);
  boxRange.setBorder(
    true,   // top
    true,   // left
    true,   // bottom
    true,   // right
    false,  // no inner vertical
    false   // no inner horizontal
  );


  SpreadsheetApp.flush();
}

function generateStripedRows(rowCount, color1, color2, colCount) {
  const result = [];
  for (let i = 0; i < rowCount; i++) {
    result.push(new Array(colCount).fill(i % 2 === 0 ? color1 : color2));
  }
  return result;
}
