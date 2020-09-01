const fs = require("fs");
const Excel = require("exceljs");
const remote = require("electron").remote;
const { dialog } = remote;

const LIST = [
  "Flexible Talent Solution",
  "Business Agility",
  "Talent Grant",
  "Core-to-Periheral",
  "Intentionally Optimistic",
  '"Corporate Jargon"',
  "Future Workforce",
  "Scale Resources",
  "Innovation",
  "Back-to-Better ",
  "COVID",
  "Technology-Driven",
  "Presenter Sneezes",
  "Flexibility",
  "Romote-First",
  "Adaptability",
  "Unlock Potential",
  "Hybrid Models",
  "Unprecedented Opportunity",
  "Redesign Work",
  "Remote Work Strategy",
  "Transformation",
  "Proven Talent",
  "New World of Work",
];

function shuffle(list) {
  let array = [...list];

  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * i);
    const temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }

  return array;
}

function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false;
  for (var i = arr1.length; i--; ) {
    if (arr1[i] !== arr2[i]) return false;
  }

  return true;
}

function testBingo(list, count) {
  alert(count);
  alert(list);
}

function runBingo(textlist, count) {
  const list = textlist.split("\n");

  const finalList = [];

  for (let i = 0; i < count; i++) {
    console.log(`Starting iteration ${i + 1}`);
    let done = false;
    let newShuffle = [];

    while (!done) {
      console.log("Checking");
      done = true;
      newShuffle = shuffle(list);

      for (const prevShuffle of finalList) {
        if (done) {
          done = !arraysEqual(newShuffle, prevShuffle);
        } else {
          console.log("Found a duplicate!");
          break;
        }
      }
    }

    console.log("Great! Adding it to the list");
    finalList.push(newShuffle);
  }

  console.log("Done. Writing file");

  const workbook = new Excel.Workbook();

  workbook.creator = "Me";
  workbook.lastModifiedBy = "Her";
  workbook.created = new Date();
  workbook.modified = new Date();

  const sheet = workbook.addWorksheet("Bingo Cards");
  sheet.addRows(finalList);

  loc = dialog.showSaveDialogSync(remote.getCurrentWindow(), {
    title: "Save Bingo List",
    defaultPath: "bingo.xlsx"
  });

  if (typeof loc === "string") workbook.xlsx.writeFile(loc);
}
