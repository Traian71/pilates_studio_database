function verificaToateSheeturile() {
  const sheeturi = {
    "Inscrieri": { functie: handleInscriere, colVerificare: 6 },
    "Programari": { functie: handleProgramare, colVerificare: 6 },
    "Programari2": { functie: handleProgramare2, colVerificare: 8 },
    "Programari3": { functie: handleProgramare3, colVerificare: 10 },
    "Reprogramari": { functie: reprogramarePilates, colVerificare: 7 },
    "Reinnoire": { functie: reinnoireAbonament, colVerificare: 3}
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const [numeSheet, { functie, colVerificare }] of Object.entries(sheeturi)) {
    const sheet = ss.getSheetByName(numeSheet);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    const lastRow = data.length;

    for (let i = lastRow - 1; i >= 1; i--) { // începem de la ultimul rând și ignorăm header-ul (rândul 0)
      const verificare = data[i][colVerificare];
      if (verificare === "verificat") {
        break; // dacă găsește un rând deja verificat, se oprește
      }

      const fakeEvent = {
        changeType: 'INSERT ROW',
        source: ss,
        range: sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()),
        sheet: sheet,
        row: i + 1
      };

      functie(fakeEvent); // apelăm funcția corespunzătoare

      sheet.getRange(i + 1, colVerificare + 1).setValue("verificat"); // marcare ca verificat
    }
  }
}


function reinnoireAbonament(e) {
  const sheet = e.sheet;
  const lastRow = sheet.getRange(e.row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const [_, prenume, nume] = lastRow;

  const abonamenteSheet = e.source.getSheetByName("Abonamente");
  const abonamenteData = abonamenteSheet.getDataRange().getValues();

  for (let i = 1; i < abonamenteData.length; i++) { // începem de la 1 ca să sărim peste antet
    const [aPrenume, aNume, , tipAbonament, , sesiuniTotale] = abonamenteData[i];

    if (aPrenume === prenume && aNume === nume) {
      const sesiuniNoi = {
        "Align": 4,
        "Elevate": 8,
        "Empower": 12,
        "Mama&Fiica": 4
      }[tipAbonament] || 0;

      const totalNou = (Number(sesiuniTotale) || 0) + sesiuniNoi;
      abonamenteSheet.getRange(i + 1, 6).setValue(totalNou); // coloana 6 = Sesiuni totale

      Logger.log(`Abonament reînnoit pentru ${prenume} ${nume}: +${sesiuniNoi} sesiuni (total: ${totalNou})`);
      return;
    }
  }

  Logger.log(`Clientul ${prenume} ${nume} nu a fost găsit în foaia 'Abonamente'.`);
}


function handleInscriere(e) {
  const sheet = e.sheet;
  const lastRow = sheet.getRange(e.row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const [_, prenume, nume, email, abonament] = lastRow;

  const dataActivarii = new Date();
  const abonamenteSheet = e.source.getSheetByName("Abonamente");
  const abonamenteData = abonamenteSheet.getDataRange().getValues();

  const sesiuni = {
    "Align": 4,
    "Elevate": 8,
    "Empower": 12,
    "Mama&Fiica": 4
  }[abonament] || 0;

  const emailExists = abonamenteData.some(row => row[2] === email);

  // Încearcă trimiterea emailului - dacă e invalid, iese din funcție
  try {
    MailApp.sendEmail(email, "Confirmare programare", 
      `Salut, te-ai înscris cu abonamentul ${abonament} la Balance Studio.`);
  } catch (e) {
    if (e.message.includes("no recipient")) {
      const emailAdmin = Session.getActiveUser().getEmail(); 
      MailApp.sendEmail(emailAdmin,
        "Eroare email client Pilates", 
        `Clientul ${prenume} ${nume} are un email invalid: ${email}.`);
      Logger.log("Email invalid. Programarea a fost anulată.");
      return; // Ieși din funcție, nu continua
    } else {
      throw e;
    }
  }

  // Dacă emailul nu există și e valid (a trecut de try), îl adăugăm
  if (!emailExists) {
    abonamenteSheet.appendRow([prenume, nume, email, abonament, dataActivarii, sesiuni]);
    Logger.log(`Client nou adăugat: ${prenume} ${nume}, ${email}, ${abonament}, ${dataActivarii}, ${sesiuni}`);
  } else {
    Logger.log(`Clientul deja există: ${prenume} ${nume}`);
  }
}



function programarePilates(e, numarZilePeSaptamana, sesiuniNecesare, coloaneDate) {
  const calendarID = 'e6fddbbafb21a03b956c106f83e570f831d8534a23113324e9c4c4f82f4aa26e@group.calendar.google.com';
  const maxLocuri = 4;
  const currentYear = new Date().getFullYear();

  const sheet = e.sheet;
  const row = e.range.getRow();
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const prenume = data[1];
  const nume = data[2];
  const nivelExperienta = data[3];

  const dateZile = [];
  const intervaleZile = [];

  for (let i = 0; i < numarZilePeSaptamana; i++) {
    dateZile.push(new Date(data[coloaneDate[i * 2]]));
    intervaleZile.push(data[coloaneDate[i * 2 + 1]]);
    dateZile[i].setFullYear(currentYear);
  }

  const abonamenteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Abonamente");
  const abonamente = abonamenteSheet.getDataRange().getValues();
  let sesiuniRamase = 0;
  let emailClient = "";

  for (let i = 1; i < abonamente.length; i++) {
    if (abonamente[i][0] === prenume && abonamente[i][1] === nume) {
      sesiuniRamase = abonamente[i][5];
      emailClient = abonamente[i][2];
      break;
    }
  }
  Logger.log("Trimit email către: " + emailClient);

  if (sesiuniRamase < sesiuniNecesare) {
    MailApp.sendEmail(emailClient, "Programare refuzată",
      `Ai nevoie de cel puțin ${sesiuniNecesare} sesiuni disponibile pentru a te programa de ${numarZilePeSaptamana} ori pe săptămână timp de 4 săptămâni.`);
    return;
  }

  const calendar = CalendarApp.getCalendarById(calendarID);
  const confirmari = [];

  for (let saptamana = 0; saptamana < 4; saptamana++) {
    for (let zi = 0; zi < numarZilePeSaptamana; zi++) {
      const interval = intervaleZile[zi];
      if (!interval || typeof interval !== 'string') continue;

      const [oraStart, oraSfarsit] = interval.split("-").map(t => t.trim());

      const dataCurenta = new Date(dateZile[zi]);
      dataCurenta.setDate(dataCurenta.getDate() + 7 * saptamana);

      const startDate = new Date(dataCurenta);
      const [startHour, startMinute] = oraStart.split(":");
      startDate.setHours(Number(startHour), Number(startMinute));

      const endDate = new Date(dataCurenta);
      const [endHour, endMinute] = oraSfarsit.split(":");
      endDate.setHours(Number(endHour), Number(endMinute));

      const events = calendar.getEvents(startDate, endDate);
      let inscrisi = [];
      let nivelEveniment = nivelExperienta;
      
      if (events.length > 0) {
        const ev = events[0];
        const descriere = ev.getDescription();
        const descriereLinii = descriere.split("\n");
        const linieNivel = descriereLinii[0];

        if (linieNivel.startsWith("Nivel:")) {
          nivelEveniment = linieNivel.split(":")[1].trim();
          if (nivelEveniment !== nivelExperienta) {
            MailApp.sendEmail(emailClient, "Programare refuzată",
              `Intervalul ales (${interval} - ${dataCurenta.toLocaleDateString()}) este rezervat pentru nivelul ${nivelEveniment}.`);
            return;
          }

          inscrisi = descriereLinii.slice(1);
        }
      }

      if (inscrisi.length >= maxLocuri) {
        MailApp.sendEmail(emailClient, "Programare Pilates – Interval Ocupat",
          `Ne pare rău, intervalul din ${dataCurenta.toLocaleDateString()} este deja ocupat. Încearcă altă oră.`);
        return;
      }

      inscrisi.push(`${prenume} ${nume}`);
      const titlu = `${interval} (${inscrisi.length}/${maxLocuri})`;
      const descriereNoua = `Nivel: ${nivelEveniment}\n${inscrisi.join("\n")}`;


      if (events.length === 0) {
      const event = calendar.createEvent(titlu, startDate, endDate, { description: descriereNoua });

      // Setează culoarea în funcție de nivelul de experiență
      let eventColor;
      if (nivelEveniment === "Incepator") {
        eventColor = CalendarApp.EventColor.CYAN;
      } else if (nivelEveniment === "Avansat") {
        eventColor = CalendarApp.EventColor.BLUE;
      } 
      event.setColor(eventColor); 
      } else {
        const event = events[0];
        event.setTitle(titlu);
        event.setDescription(descriereNoua);
        Logger.log(`Numar inscrisi: ${inscrisi.length}`);
        if (inscrisi.length === maxLocuri) {
          event.setColor(CalendarApp.EventColor.RED);
        }

      }

      for (let i = 1; i < abonamente.length; i++) {
        if (abonamente[i][0] === prenume && abonamente[i][1] === nume) {
          abonamente[i][5]--;
          abonamenteSheet.getRange(i + 1, 6).setValue(abonamente[i][5]);
          break;
        }
      }

      confirmari.push(`✔️ ${dataCurenta.toLocaleDateString()} la ora ${interval}`);
    }
  }

  MailApp.sendEmail(emailClient, "Programare confirmată – Pilates",
    `Te-ai programat cu succes la Pilates în următoarele date:\n\n${confirmari.join("\n")}`);
}


function reprogramarePilates(e) {
  const calendarID = 'e6fddbbafb21a03b956c106f83e570f831d8534a23113324e9c4c4f82f4aa26e@group.calendar.google.com';
  const maxLocuri = 4;
  const calendar = CalendarApp.getCalendarById(calendarID);

  const sheet = e.sheet;
  const row = e.range.getRow();
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const prenume = data[1];
  const nume = data[2];
  const dataInitiala = new Date(data[3]);
  const intervalInitial = data[4];
  const dataNoua = new Date(data[5]);
  const intervalNou = data[6];
  const currentYear = new Date().getFullYear();

  // === Detalii eveniment inițial ===
  const [oraStartInit, oraSfarsitInit] = intervalInitial.split("-").map(t => t.trim());
  const startInit = new Date(dataInitiala);
  const endInit = new Date(dataInitiala);
  startInit.setFullYear(currentYear);
  endInit.setFullYear(currentYear);
  startInit.setHours(...oraStartInit.split(":").map(Number));
  endInit.setHours(...oraSfarsitInit.split(":").map(Number));

  const evenimenteInitiale = calendar.getEvents(startInit, endInit);
  if (evenimenteInitiale.length === 0) return;

  const evenimentInitial = evenimenteInitiale[0];
  const descriereInitiala = evenimentInitial.getDescription();
  const descriereLinii = descriereInitiala.split("\n");
  const linieNivel = descriereLinii[0].replace(/\r/g, "").trim();
  if (!linieNivel.startsWith("Nivel:")) return;
  const nivel = linieNivel.split(":")[1].trim();

  let inscrisiInitial = descriereLinii.slice(1).map(l => l.trim()).filter(n => n !== `${prenume} ${nume}`);

  // === Găsim emailul clientului ===
  const abonamenteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Abonamente");
  const abonamente = abonamenteSheet.getDataRange().getValues();
  let emailClient = "";
  for (let i = 1; i < abonamente.length; i++) {
    if (abonamente[i][0] === prenume && abonamente[i][1] === nume) {
      emailClient = abonamente[i][2];
      break;
    }
  }

  // === Verificăm dacă noul interval este OK ===
  const [oraStartNou, oraSfarsitNou] = intervalNou.split("-").map(t => t.trim());
  const startNou = new Date(dataNoua);
  const endNou = new Date(dataNoua);
  startNou.setFullYear(currentYear);
  endNou.setFullYear(currentYear);
  startNou.setHours(...oraStartNou.split(":").map(Number));
  endNou.setHours(...oraSfarsitNou.split(":").map(Number));

  const evenimenteNoi = calendar.getEvents(startNou, endNou);
  let reprogramareReusita = false;

  if (evenimenteNoi.length > 0) {
    const ev = evenimenteNoi[0];
    const desc = ev.getDescription();
    const linii = desc.split("\n");
    if (linii[0].startsWith("Nivel:")) {
      const nivelNou = linii[0].split(":")[1].trim();
      if (nivelNou !== nivel) {
        MailApp.sendEmail(emailClient, "Reprogramare refuzată",
          `Intervalul ales (${intervalNou} - ${dataNoua.toLocaleDateString()}) este rezervat pentru nivelul ${nivelNou}.`);
        return;
      }

      let inscrisi = linii.slice(1).map(l => l.trim());
      if (inscrisi.length >= maxLocuri) {
        MailApp.sendEmail(emailClient, "Reprogramare Pilates – Interval Ocupat",
          `Ne pare rău, intervalul din ${dataNoua.toLocaleDateString()} este deja ocupat. Încearcă altă oră.`);
        return;
      }

      inscrisi.push(`${prenume} ${nume}`);
      ev.setDescription(`Nivel: ${nivelNou}\n${inscrisi.join("\n")}`);
      ev.setTitle(`${intervalNou} (${inscrisi.length}/${maxLocuri})`);
      if (inscrisi.length === maxLocuri) ev.setColor(CalendarApp.EventColor.RED);
      reprogramareReusita = true;
    }
  } else {
    const descriereNoua = `Nivel: ${nivel}\n${prenume} ${nume}`;
    const titluNou = `${intervalNou} (1/${maxLocuri})`;
    const ev = calendar.createEvent(titluNou, startNou, endNou, { description: descriereNoua });

    let eventColor = nivel === "Incepator" ? CalendarApp.EventColor.CYAN : CalendarApp.EventColor.BLUE;
    ev.setColor(eventColor);
    reprogramareReusita = true;
  }

  // === Dacă reprogramarea a fost cu succes, modificăm/suprimăm vechiul eveniment ===
  if (reprogramareReusita) {
    if (inscrisiInitial.length === 0) {
      evenimentInitial.deleteEvent();
    } else {
      evenimentInitial.setDescription(`Nivel: ${nivel}\n${inscrisiInitial.join("\n")}`);
      evenimentInitial.setTitle(`${intervalInitial} (${inscrisiInitial.length}/${maxLocuri})`);
    }

    MailApp.sendEmail(emailClient, "Reprogramare confirmată – Pilates",
      `Te-ai reprogramat cu succes la Pilates:\n\n✔️ ${dataNoua.toLocaleDateString()} la ora ${intervalNou}`);
  }
}



function handleProgramare(e) {
  programarePilates(e, 1, 4, [4, 5]);
}

function handleProgramare2(e) {
  programarePilates(e, 2, 8, [4, 5, 6, 7]);
}

function handleProgramare3(e) {
  programarePilates(e, 3, 12, [4, 5, 6, 7, 8, 9]);
}
