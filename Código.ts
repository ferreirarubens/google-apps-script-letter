var CONST_LETTER_FOLDER_ID = "1GjPisC6GY_cHDoBpGRJXkX3gt7Ukedvt";
var CONST_MUSICS_FOLDER_ID = "1xRCnyVrpvnbzEaLeiW_DNSv4QIhTXyjw";
var CONST_MODEL_ID = "1_b0Ub9PbIqNk9P34Y1gzOxr39Au_LAkgpCNUmnjcqpo";
var VAGALUME_API_KEY = "";
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
  artist: string;
  letter: string;
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
}

function testCreationMusic() {
  var music = {
    title: "NINGUÉM SE IMPORTA COM MINHA ALMA",
    artist: "JOSUÉ LIRA",
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

function executeLetterScriptFromSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("MUSIC"), true);

  const initialRangeIndex = 2;

  // get all range of music
  var range = spreadsheet.getRange(`A${initialRangeIndex}:E`);
  var values = range.getValues();

  checkRelatedArtists(values);

  // GET ROWs from range IF A1 is CHECKED
  var rows = values.filter((value) => value[0] === true && value[2] !== "");
  total = rows.length;
  // iterate rows and transform into music object
  var musics = rows.map(
    (value) =>
      ({
        title: value[2],
        artist: value[3],
        id: value[1],
        referenceArtistsLetter: value[4],
      } as Music)
  );

  // log the first music
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
  const range = sheet.getRange(`A2:G`);
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
      .replace('POUT PORRI - ', '[remove]')
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
          docName = `${music.referenceArtistsLetter ? '[referenced]' : ''} ${music.title} - ${music.artist}`;
          break;
      }
    

      var copyId = DriveApp.getFileById(CONST_MODEL_ID)
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

      var saveFolder = DriveApp.getFolderById(CONST_LETTER_FOLDER_ID);
      saveFolder.addFile(DriveApp.getFileById(copyId));
      copyDoc.saveAndClose();
      music.idDocument = copyId;
      updateSpreadsheetWithStatusAndId(music);
      // toPdf(saveFolder.getId(), copyId);

      if (count === total) {
        openFolder(saveFolder.getId());
      }
    });
}

function openFolder(id) {
  showURL("https://drive.google.com/drive/u/0/folders/".concat(id));
}

function showURL(href) {
  var html = '\n  <html>\n    <body>\n      <a href="'.concat(
    href,
    '" target="blank" onclick="google.script.host.close()">Abrir pasta das letras</a>\n      </body>\n      </html>\n      '
  );
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
  menu.addItem("Gerar Letras", "executeLetterScriptFromSpreadsheet");
  menu.addItem("Atualizar Músicas", "updateSpreadsheetWithAllMusicsFromFolder");
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
      music.letter = returnObj.mus[0].text;
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
  const titleAndArtistWithoutAccents = removeAccent(
    music.title + " - " + music.referenceArtistsLetter
      ? music.referenceArtistsLetter
      : music.artist
  );

  Logger.log(titleAndArtistWithoutAccents);
  music.logger = music.logger + "\n" + titleAndArtistWithoutAccents;

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
      removeAccent(music.referenceArtistsLetter)
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

      music.logger =
        music.logger +
        "\n" +
        "Most related- " +
        JSON.stringify(mostRelatedTitle);
      music.based = `${mostRelatedTitle.mus} - ${mostRelatedTitle.artist}`;

      idLetter = returnObj.response.docs.find(
        (docs) =>
          removeAccent(docs.band.toUpperCase()) === mostRelatedTitle.artist
      ).id;
      music.logger = music.logger + "\n" + "idLetter- " + idLetter;
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
  music.letter = returnObj2.mus[0].text;
  return music;
}

function generateQueryString(data) {
  const params = [];
  for (var d in data)
    params.push(encodeURIComponent(d) + "=" + encodeURIComponent(data[d]));
  return params.join("&");
}
