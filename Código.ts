var CONST_LETTER_FOLDER_DOC_ID = "1gjpisc6gy_chdobpgrjxkx3gt7ukedvt";
var CONST_LETTER_FOLDER_PDF_ID = "198lRqvb2ZZl7X6p_rLSqzqDgEr1T0cOS";
var CONST_MUSICS_FOLDER_ID = "1xrcnyvrpvnbzealeiw_dnsv4qihtxyjw";
var CONST_MODEL_ID_DOC_PDF = "1_b0Ub9PbIqNk9P34Y1gzOxr39Au_LAkgpCNUmnjcqpo";
var CONST_MODEL_ID_REVIEW_DOC = "19NWovX0yHBA0zWrAd6S1tYRlDN6Aiv_OnOE7UrZ8blY";
var CONST_FOLDER_LETTERS = "18lhEbOvYgtZSNQHZBWcv3ih06iGyaaOb";
const maxLine = 30;

const VAGALUME_API_KEY = "384fe401b98e39a9f44ebfc6479df5d7";
var vars = ["title", "artist"];
var relatedArtists = new Set<string>();
/* var relatedArtists = [
  "CARLOS JOSE",
  "GERSON RUFINO",
  "NANI AZEVEDO",
  "WELLINGTON DEMEZIO",
  "CICERO NOGUEIRA",
  "IVONALDO ALBUQUERQUE",
  "VICTORINO SILVA",
  "OSEIAS O SEMEADOR",
  "SAULO NOGUEIRA",
  "OZEIAS DE PAULA",
  "JUNIOR",
  "ESTEVES JACINTO",
  "CHAGAS SOBRINHO",
  "ALCEU PIRES",
  "DANIEL E SAMUEL",
  "ELIEZER ROSA",
  "THIAGO PAZ",
  "JONATHAS ALMEIDA",
  "JAIR PIRES",
  "EDGAR MARTINS",
  "ADILSON LOPES",
  "LUIZ DE CARVALHO",
  "PAULO FIGUEIREDO",
  "ISAC SA",
  "MATTOS NASCIMENTO",
  "EDUARDO SILVA",
  "GESIEL MENDES",
  "ROBINSON MONTEIRO",
  "ESTEVAO SILVA",
  "JAMENSON LUIZ",
  "IRISVALDO SILVA",
  "OS LEVITAS",
  "SILVAN SANTOS",
  "JOSE TOSTES",
  "VOZ DA VERDADE",
  "PR. MOISES DE OLIVEIRA",
  "JOSE CARLOS",
  "RUBEM LOPES",
  "JOSUE LIRA",
  "KALEBE",
  "ADILSON ROSSI",
  "OSVALDO NASCIMENTO",
  "MARILENE SANTIAGO",
  "JOSE REIS",
  "MARCOS SILVA",
  "EZEQUIAS OLIVEIRA",
  "ARMANDO FILHO",
  "EMBAIXADORES DE SIAO",
  "MISSIONARIO DUARTE",
  "DEBORA IVANOV",
  "SHIRLEY CARVALHAES",
  "DEUS NAO PERDE EM QUESTAO",
  "JA O VERBO ERA DEUS",
  "QUEBRANTA-ME",
  "VAI ME AJUDAR",
  "GRUPO NOVA DIMENSAO",
  "RAURYCELIA",
  "DESCONHECIDO",
  "MARCO AURELIO",
  "CANCAO E LOUVOR",
  "VOCAL",
]; */

var total = 0;
var count = 0;

interface Music {
  title: string;
  artist?: string;
  letter?: string;
  idDocument?: string;
  idPdf?: string;
  status?:
    | "not_found"
    | "not_trusted"
    | "30_trusted"
    | "60_trusted"
    | "70_trusted"
    | "trusted"
    | "created"
    | "referenced"
    | "imported";
  indexRow?: number;
  referenceArtistsLetter?: string;
  id?: string;
  based?: string;
  logger?: string;
  letterLinesCount?: number;
}

function testCreationMusic() {
  var music = {
    title: "JUÍZO FINAL",
    artist: "VICTORINO SILVA",
    letter: "",
  } as Music;
  createMusic(music);
}

function checkRelatedArtists(values) {
  values
    .filter((row) => row[3] != "")
    .forEach((row) => {
      relatedArtists.add(removeAccent(row[3].trim().toUpperCase()));
    });
}

function executeReviewAndGenerate(musics: Music[]) {
  // 3 - get body of docs
  musics = musics.filter((el) => el != null || el != undefined);
  var bodies = musics.map((el) => {
    var doc = DocumentApp.openById(el.idDocument);
    return doc.getBody();
  });

  // 4 - join body of docs in one
  var newDocId = DriveApp.getFileById(CONST_MODEL_ID_REVIEW_DOC)
    .makeCopy(`${musics.map((el) => el.title).join("/")}`)
    .getId();

  // Open the temporary document
  var copyDoc = DocumentApp.openById(newDocId);
  var copyBody = (copyDoc as any).getActiveSection();
  // map the body of the docs
  const lettersOnly = bodies.map((el) => {
    var arr = el.getText().split("\n");
    //remove line with 'música baseada na versão'
    arr = arr.filter((el) => !el.includes("música baseada na versão"));
    // get slice from the 3rd line (letter)
    return arr.slice(3).join("\n").trim();
  });

  for (var i = 0; i < musics.length; i++) {
    copyBody.replaceText(`{{title${i + 1}}}`, musics[i].title);
    copyBody.replaceText(`{{artist${i + 1}}}`, musics[i].artist);
    //if letterLinesCount is less than maxLine concat with difference of maxLine
    var letter =
      musics[i].letterLinesCount < maxLine
        ? lettersOnly[i].concat(
            Array(maxLine - musics[i].letterLinesCount).join("\n")
          )
        : lettersOnly[i];

    copyBody.replaceText(
      `{{letter${i + 1}}}`,
      i === 0 ? letter : letter.trim()
    );
  }

  if (musics.length === 1) {
    copyBody.replaceText(`{{title2}}`, "");
    copyBody.replaceText(`{{artist2}}`, "");
    copyBody.replaceText(`{{letter2}}`, "");
  }

  // get folder by id
  var saveFolder = DriveApp.getFolderById(`${CONST_FOLDER_LETTERS}`);
  saveFolder.addFile(DriveApp.getFileById(newDocId));
  copyDoc.saveAndClose();

  return newDocId;
}

function reviewAndGenerate() {
  // 1 - select from spreadsheet the music list who have checked(A)
  var musics = getRowsFromSpreadsheetWhoHaveChecked("I").map((el) => {
    // 2 - get the ID Doc (G)
    return {
      title: el[2],
      artist: el[3],
      idDocument: el[6],
      letterLinesCount: el[8],
    } as Music;
  });

  const urlsFromMusics = [];

  // sort desc by letterLinesCount
  musics = musics.sort((a, b) => {
    return b.letterLinesCount - a.letterLinesCount;
  });

  // if letterLinesCount is less than maxLine, then the music comes first
  musics = musics
    .filter((el) => el.letterLinesCount <= maxLine)
    .concat(musics.filter((el) => el.letterLinesCount > maxLine));

  // execute review and generate two musics by 2 musics
  for (let i = 0; i < musics.length; i += 2) {
    const newDocId = executeReviewAndGenerate([musics[i], musics[i + 1]]);
    urlsFromMusics.push({
      titles:
        musics.length > 1
          ? [musics[i].title, musics[i + 1].title].join("/")
          : musics[i].title,
      url: `https://docs.google.com/document/d/${newDocId}/edit`,
    });
  }

  showURLReview(urlsFromMusics);
}

function showURLReview(urlsFromMusics) {
  var html = `
    <html>
      <body>
      ${urlsFromMusics
        .map((el) => `<a href="${el.url}" target="blank">${el.titles}</a>`)
        .join("<br>")}
      </body>
    </html>
    `;
  var ui = HtmlService.createHtmlOutput(html);
  ui.setWidth(500);
  ui.setHeight(urlsFromMusics.length * 30);
  SpreadsheetApp.getUi().showModelessDialog(ui, "Letras geradas com sucesso!");
}

function getRowsFromSpreadsheetWhoHaveChecked(rangeEnd) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("MUSIC"), true);

  const initialRangeIndex = 2;

  // get all range of music
  var range = spreadsheet.getRange(`A${initialRangeIndex}:${rangeEnd}`);
  var values = range.getValues();

  checkRelatedArtists(values);

  // GET ROWs from range IF A1 is CHECKED
  var rows = values.filter((value) => value[0] === true && value[2] !== "");
  total = rows.length;
  // iterate rows and transform into music object
  return rows;
}

function executeLetterScriptFromSpreadsheet() {
  var musics = getRowsFromSpreadsheetWhoHaveChecked("E").map(
    (value) =>
      ({
        title: value[2],
        artist: value[3],
        id: value[1],
        referenceArtistsLetter: value[4],
      } as Music)
  );
  musics.forEach((music) => createMusic(music));
}

function generateRandomId() {
  return Math.random().toString(36).substring(2, 15);
}

function testAndReplace() {
  const music = {} as Music;
  updateSpreadsheetWithStatusAndId({ ...music, id: "ctdh0o0p61p" });
}

function updateSpreadsheetWithStatusAndId(music: Music) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("MUSIC");

  // find cell by value
  const range = sheet.getRange(`A2:I`);
  const data = range.getValues();

  var indexRow = 0;

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == music.id) {
      indexRow = i + 2;
      break;
    }
  }

  // update status
  sheet.getRange(`A${indexRow}`).setValue(false);
  sheet.getRange(`F${indexRow}`).setValue(music.status);
  sheet.getRange(`G${indexRow}`).setValue(music.idDocument);
  sheet.getRange(`H${indexRow}`).setValue(music.idPdf);
  sheet.getRange(`I${indexRow}`).setValue(music.letter.split("\n").length);
}

function updateSpreadsheetWithAllMusicsFromFolder() {
  const musicsFolder = DriveApp.getFolderById(CONST_MUSICS_FOLDER_ID);

  const musicsFiles = musicsFolder.getFiles();
  // get file names
  const setMusicsNames = new Set();
  while (musicsFiles.hasNext()) {
    const music = musicsFiles.next();
    // split title and artist
    // remove parenthesis, content in parenthesis and double spaces, and extension mp3 and remove prefix "HC d{3} - "
    const withoutParenthesis = music
      .getName()
      .replace(/\(.*\)/g, "")
      .replace(/\s\s+/g, " ")
      .replace(/\.mp3/g, "")
      .replace(/^HC \d{3} - /g, "[remove]")
      .replace(/^HC \d{2} - /g, "[remove]")
      .replace(/^HC \d{1} - /g, "[remove]")
      .replace("POUT PORRI - ", "[remove]")
      .toUpperCase();
    setMusicsNames.add(withoutParenthesis);
  }

  // map set to Music array
  const musicsNames = (Array.from(setMusicsNames) as string[])
    .filter((music) => !music.startsWith("[REMOVE]"))
    .map((music) => {
      const [title, artist] = music.split(" - ");
      return {
        title,
        artist,
        id: generateRandomId(),
        status: "imported",
      };
    });

  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("MUSIC");
  const range = sheet.getRange(`B2:F${musicsNames.length + 1}`);
  // map musics to range values
  const values = musicsNames.map((music) => [
    music.id,
    music.title,
    music.artist,
    "",
    music.status,
  ]);
  range.setValues(values);
}

function createMusic(music: Music) {
  try {
    // get the letter of music
    getLettersByArtistAndTitle(music)
      .then((musicResp: Music) => {
        music = musicResp;
      })
      .catch((error) => {
        music = getLettersByTitle(music);
      })
      .finally(() => {
        count++;

        var docName = "";

        switch (music.status) {
          case "referenced":
            docName = `[referenced] ${
              music.title
            } - ${music.referenceArtistsLetter.toUpperCase()}`;
            break;
          case "not_trusted":
            docName = `[based] ${music.title} - ${music.artist}`;
            break;
          default:
            docName = `${music.referenceArtistsLetter ? "[referenced]" : ""} ${
              music.title
            } - ${music.artist}`;
            break;
        }

        var copyId = DriveApp.getFileById(CONST_MODEL_ID_DOC_PDF)
          .makeCopy(`${docName}`)
          .getId();

        // Open the temporary document
        var copyDoc = DocumentApp.openById(copyId);
        // Get the document’s body section
        var copyBody = (copyDoc as any).getActiveSection();

        copyBody.replaceText(
          "{{based}}",
          music.based ? `música baseada na versão ${music.based}` : ""
        );
        // copyBody.replaceText("{{logger}}", "");
        copyBody.replaceText("{{title}}", music.title);
        copyBody.replaceText("{{artist}}", music.artist);
        copyBody.replaceText("{{letter}}", music.letter);

        // get folder by id
        var saveFolder = DriveApp.getFolderById(
          "1GjPisC6GY_cHDoBpGRJXkX3gt7Ukedvt"
        );
        Logger.log(saveFolder.getName());
        saveFolder.addFile(DriveApp.getFileById(copyId));
        copyDoc.saveAndClose();
        music.idDocument = copyId;
        const pdf = toPdf(CONST_LETTER_FOLDER_PDF_ID, copyId);
        music.idPdf = pdf.pdfId;
        updateSpreadsheetWithStatusAndId(music);

        if (count === total) {
          showURL(saveFolder.getId(), pdf.pdfFolderId);
        }
      });
  } catch (error) {
    Logger.log(error);
  }
}

function showURL(hrefDOC, hrefPDF) {
  var html = `
    <html>
      <body>
        <a href="https://drive.google.com/drive/u/0/folders/${hrefDOC}" target="blank" onclick="google.script.host.close()">[DOC] Abrir pasta das letras</a>
        <br>
        <a href="https://drive.google.com/drive/u/0/folders/${hrefPDF}" target="blank" onclick="google.script.host.close()">[PDF] Abrir pasta das letras</a>
      </body>
    </html>
    `;
  var ui = HtmlService.createHtmlOutput(html);
  ui.setWidth(300);
  ui.setHeight(100);
  SpreadsheetApp.getUi().showModelessDialog(ui, "Letras geradas com sucesso!");
}

var toPdf = function (folderID, docId) {
  // create PDF version of doc *************************************************
  var docFile = DriveApp.getFileById(docId);
  var blobFile = docFile.getAs("application/pdf");
  var pdfVersion = DriveApp.createFile(blobFile);
  // get ID of PDF version
  var pdfVersionID = pdfVersion.getId();
  // get PDF file
  var pdfFile = DriveApp.getFileById(pdfVersionID);
  // get location of where PDF file currently exists (so 'My Drive')
  var parents = pdfFile.getParents();
  /* add PDF file to new folder location (so is ignored from above 'parents'
    and hence will not be removed from this new location) */
  DriveApp.getFolderById(folderID).addFile(pdfVersion);
  /* once PDF file has moved, remove it from any other locations so it
    only exists in the new location */
  while (parents.hasNext()) {
    var parent = parents.next();
    parent.removeFile(pdfFile);
  }

  return { pdfId: pdfVersionID, pdfFolderId: folderID };
  // delete the original Google doc file
  // docFile.setTrashed(true);
};

function onOpen() {
  addMenus();
}

function addMenus() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  var menu = ui.createMenu("Músicas");
  menu.addItem("Revisar e preparar para imprimir", "reviewAndGenerate");
  menu.addItem("Gerar Letras", "executeLetterScriptFromSpreadsheet");
  // menu.addItem("Atualizar Músicas", "updateSpreadsheetWithAllMusicsFromFolder");
  menu.addToUi();
}

function getLettersByArtistAndTitle(music: Music) {
  return new Promise<Music>((resolve, reject) => {
    try {
      var data = {
        art: music.referenceArtistsLetter
          ? music.referenceArtistsLetter
          : music.artist,
        mus: music.title,
        apiKey: VAGALUME_API_KEY,
      };

      var response = UrlFetchApp.fetch(
        "https://api.vagalume.com.br/search.php?" + generateQueryString(data)
      );
      const returnObj = JSON.parse(response.getContentText("UTF-8"));
      music.letter = returnObj.mus[0].text.replace("\n\n", "\n");
      music.status = "trusted";

      if (music.referenceArtistsLetter) {
        music.based = music.referenceArtistsLetter;
      }

      resolve(music);
    } catch (e) {
      reject(e);
    }
  });
}

function removeAccent(text: string) {
  return (
    text
      //lowercase
      .replace(/[áàãâä]/g, "a")
      .replace(/[éèêë]/g, "e")
      .replace(/[íìîï]/g, "i")
      .replace(/[óòõôö]/g, "o")
      .replace(/[úùûü]/g, "u")
      .replace(/[ç]/g, "c")
      .replace(/[ñ]/g, "n")
      //uppercase
      .replace(/[ÁÀÃÂÄ]/g, "A")
      .replace(/[ÉÈÊË]/g, "E")
      .replace(/[ÍÌÎÏ]/g, "I")
      .replace(/[ÓÒÕÔÖ]/g, "O")
      .replace(/[ÚÙÛÜ]/g, "U")
      .replace(/[Ç]/g, "C")
      .replace(/[Ñ]/g, "N")
  );
}

function getLettersByTitle(music: Music) {
  //remove accents
  const artistSearch = music.referenceArtistsLetter
    ? music.referenceArtistsLetter
    : music.artist;
  const titleAndArtistWithoutAccents = removeAccent(
    music.title +
      " - " +
      (music.referenceArtistsLetter
        ? music.referenceArtistsLetter
        : music.artist)
  );

  Logger.log(titleAndArtistWithoutAccents);

  var data = {
    q: titleAndArtistWithoutAccents,
    limit: 10,
    apiKey: VAGALUME_API_KEY,
  } as any;

  var response = UrlFetchApp.fetch(
    "https://api.vagalume.com.br/search.excerpt?" + generateQueryString(data)
  );
  const returnObj = JSON.parse(response.getContentText("UTF-8"));

  var idLetter = returnObj.response.docs[0].id;
  music.status = "not_trusted";

  // check if reference artist is the same as the one in the letter
  const referencDocBand = returnObj.response.docs.find((doc) => {
    return (
      removeAccent(doc.band.toLowerCase()) ===
      removeAccent(
        music.referenceArtistsLetter
          ? music.referenceArtistsLetter
          : music.artist
      )
    );
  });

  if (!referencDocBand) {
    var artists = returnObj.response.docs.map((doc) => {
      const dataRA = {
        art: doc.band,
        extra: "relart",
        apiKey: VAGALUME_API_KEY,
      };
      const rA = UrlFetchApp.fetch(
        "https://api.vagalume.com.br/search.php?" + generateQueryString(dataRA)
      );

      const returnRA = JSON.parse(rA.getContentText("UTF-8"));

      const artistsRelated = {
        artist: removeAccent(doc.band).toUpperCase(),
        mus: doc.title,
        related: [],
      } as any;

      if (returnRA.type !== "notfound") {
        returnRA.art.related.forEach((el) => {
          // check if relatedArtists has artist
          if (Array.from(relatedArtists).indexOf(el.name) === -1) {
            artistsRelated.related.push(removeAccent(el.name.toUpperCase()));
          }
        });
      }

      return artistsRelated;
    });

    // get the 5 top related artists and related more than 1
    const top5Artists = artists
      .sort((a, b) => {
        return b.related.length - a.related.length;
      })
      .filter((el) => {
        return el.related.length > 1;
      })
      .slice(0, 5);

    const titleWords = removeAccent(music.title).toUpperCase().split(" ");

    if (top5Artists.length > 0) {
      const titleMostRelated_ = top5Artists
        .map((el) => {
          return removeAccent(el.mus).toUpperCase().split(" ");
        })
        .map((el) => {
          var countIncludes = 0;
          titleWords.forEach((w) => {
            if (el.includes(w)) {
              countIncludes++;
            }
          });

          return {
            phrase: el.join(" "),
            includes: countIncludes,
          } as any;
        });
      const sortedTitleMostRelated = titleMostRelated_.sort((a, b) => {
        return b.includes - a.includes;
      })[0];

      // get the id of the letter
      var mostRelatedTitle = artists.find((doc) => {
        return (
          removeAccent(doc.mus).toUpperCase() ===
          removeAccent(sortedTitleMostRelated.phrase).toUpperCase()
        );
      });

      music.based = `${mostRelatedTitle.mus} - ${mostRelatedTitle.artist}`;

      idLetter = returnObj.response.docs.find(
        (docs) =>
          removeAccent(docs.band.toUpperCase()) === mostRelatedTitle.artist
      ).id;
    } else {
      const sameTitle = returnObj.response.docs.filter(
        (doc) =>
          removeAccent(doc.title.toUpperCase()) ===
          removeAccent(music.title.toUpperCase())
      );

      if (sameTitle.length > 1) {
        music.status = "70_trusted";
        // random between 0 and sameTitle.length - 1
        const random = Math.floor(Math.random() * sameTitle.length);
        idLetter = sameTitle[random].id;
        music.based = `${sameTitle[random].title} - ${sameTitle[random].band}`;
      }

      // get the most relevant letter
      const hasWordsOnTitle = returnObj.response.docs.filter((mus) =>
        removeAccent(mus.title.toUpperCase()).includes(
          removeAccent(music.title.toUpperCase())
        )
      );

      if (sameTitle.length === 0 && hasWordsOnTitle.length > 0) {
        music.status = "30_trusted";
        const random = Math.floor(Math.random() * hasWordsOnTitle.length);
        idLetter = hasWordsOnTitle[random].id;
        music.based = `${hasWordsOnTitle[random].title} - ${hasWordsOnTitle[random].band}`;
      }

      const sameArtist = returnObj.response.docs.filter((mus) =>
        removeAccent(music.artist.toUpperCase()).includes(
          removeAccent(mus.band.toUpperCase())
        )
      );

      if (
        sameTitle.length === 0 ||
        hasWordsOnTitle.length === 0 ||
        sameArtist.length > 0
      ) {
        music.status = "60_trusted";
        const random = Math.floor(Math.random() * sameArtist.length);
        idLetter = sameArtist[random].id;
        music.based = `${sameArtist[random].title} - ${sameArtist[random].band}`;
      }
    }
  } else {
    idLetter = referencDocBand.id;
    music.status = "referenced";
    music.based = `${referencDocBand.title} - ${referencDocBand.band}`;
  }
  // get letter from id
  data = {
    musid: idLetter,
    ...{ apiKey: "" },
  };

  response = UrlFetchApp.fetch(
    "https://api.vagalume.com.br/search.php?" + generateQueryString(data)
  );

  const returnObj2 = JSON.parse(response.getContentText("UTF-8"));
  music.letter = returnObj2.mus[0].text.replace("\n\n", "\n");
  return music;
}

function generateQueryString(data) {
  const params = [];
  for (var d in data)
    params.push(encodeURIComponent(d) + "=" + encodeURIComponent(data[d]));
  return params.join("&");
}
