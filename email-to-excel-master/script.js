let convertedData = [];

function parseEmail(email) {
  email = email.trim().toLowerCase();
  if (!email.includes('@')) return null;

  const localPart = email.split('@')[0];
  let nom = '', prenom = '';

  if (localPart.includes('.')) {
    const parts = localPart.split('.');
    nom = parts[0];
    prenom = parts.slice(1).join('.');
  } else if (localPart.includes('_')) {
    const parts = localPart.split('_');
    nom = parts[0];
    prenom = parts.slice(1).join('_');
  } else {
    let bestSplit = findBestSplit(localPart);
    nom = localPart.substring(0, bestSplit);
    prenom = localPart.substring(bestSplit);
  }

  return {
    nom: capitalizeFirst(nom),
    prenom: capitalizeFirst(prenom),
    email
  };
}

function findBestSplit(text) {
  const length = text.length;
  const vowels = 'aeiou';
  const mid = Math.floor(length / 2);

  for (let i = Math.max(3, mid - 3); i <= Math.min(length - 3, mid + 3); i++) {
    if (vowels.includes(text[i]) && !vowels.includes(text[i + 1])) {
      return i + 1;
    }
  }
  return mid;
}

function capitalizeFirst(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function convertEmails() {
  const input = document.getElementById('emailInput').value;
  const emails = input.split('\n').filter(line => line.trim());
  convertedData = [];

  emails.forEach(email => {
    const parsed = parseEmail(email);
    console.log('Email analysé:', email, 'Résultat:', parsed);
    if (parsed) convertedData.push(parsed);
  });

  console.log('convertedData après analyse:', convertedData);
  displayResults();
}

function displayResults() {
  const preview = document.getElementById('preview');
  const tableBody = document.getElementById('tableBody');
  const countDisplay = document.getElementById('countDisplay');
  const downloadBtn = document.querySelector('.btn-download');
  const csvBtn = document.querySelector('.btn-csv');

  if (convertedData.length === 0) {
    preview.style.display = 'none';
    downloadBtn.disabled = true;
    downloadBtn.disabled = true;
    console.log('Aucune donnée à afficher');
    return;
  }

  countDisplay.textContent = `${convertedData.length} email(s) traité(s)`;
  tableBody.innerHTML = '';

  convertedData.forEach((data, index) => {
    const row = tableBody.insertRow();
    console.log('Ajout ligne tableau:', data);

    ['nom', 'prenom'].forEach(field => {
      const cell = row.insertCell();
      cell.textContent = data[field];
      cell.className = 'editable';
      cell.dataset.field = field;
      cell.dataset.index = index;
      cell.addEventListener('dblclick', startEdit);
    });

    const emailCell = row.insertCell();
    emailCell.textContent = data.email;
  });

  preview.style.display = 'block';
  downloadBtn.disabled = false;
  csvBtn.disabled = false;
  console.log('Affichage terminé, tableau mis à jour.');
}

function startEdit(event) {
  const cell = event.target;
  const field = cell.dataset.field;
  const index = parseInt(cell.dataset.index);
  const currentValue = cell.textContent;

  if (cell.querySelector('input')) return;

  const input = document.createElement('input');
  input.type = 'text';
  input.value = currentValue;
  cell.innerHTML = '';
  cell.appendChild(input);
  cell.classList.add('editing');

  input.focus();
  input.select();

  const saveEdit = () => {
    const newValue = input.value.trim();
    if (newValue) {
      convertedData[index][field] = capitalizeFirst(newValue);
      cell.textContent = capitalizeFirst(newValue);
    } else {
      cell.textContent = currentValue;
    }
    cell.classList.remove('editing');
  };

  input.addEventListener('blur', saveEdit);
  input.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') input.blur();
  });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      cell.textContent = currentValue;
      cell.classList.remove('editing');
    }
  });
}

function downloadExcel() {
  if (convertedData.length === 0) {
    alert("Aucune donnée à exporter.");
    return;
  }

  console.log("Début de l'export Excel");
  const emailsPerSheetValue = parseInt(document.getElementById('emailsPerSheet').value, 10);
  const totalEmails = convertedData.length;
  console.log("Nombre total d'emails:", totalEmails);
  
  // Déterminer le mode d'export
  const isAllInOneSheet = emailsPerSheetValue === 0;
  if (isAllInOneSheet) {
    console.log("Mode une seule feuille sélectionné - tous les emails seront dans une seule feuille");
  }
  
  const emailsPerSheet = isAllInOneSheet ? totalEmails : (emailsPerSheetValue || 15);
  const expectedSheets = isAllInOneSheet ? 1 : Math.ceil(totalEmails / emailsPerSheet);
  
  console.log("Nombre d'emails par feuille:", emailsPerSheet);
  console.log("Nombre de feuilles attendu:", expectedSheets);

  const wb = XLSX.utils.book_new();
  const headers = ['Nom', 'Prénom', 'Email'];

  let sheetCount = 0;
  for (let i = 0; i < totalEmails; i += emailsPerSheet) {
    const currentEmails = convertedData.slice(i, i + emailsPerSheet);
    console.log(`Création feuille ${sheetCount + 1} avec ${currentEmails.length} emails`);
    
    const sheetData = [headers, ...currentEmails.map(d => [d.nom, d.prenom, d.email])];
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws['!cols'] = [{ wch: 15 }, { wch: 15 }, { wch: 30 }];
    const sheetName = `Feuille_${sheetCount + 1}`;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    sheetCount++;
  }

  console.log(`Export terminé: ${sheetCount} feuille(s) créée(s)`);
  window.alert(`${sheetCount} feuille(s) générée(s) dans le classeur Excel`);
  
  const currentDate = new Date().toISOString().split("T")[0];
  XLSX.writeFile(wb, `contacts_${currentDate}.xlsx`);
}

function clearAll() {
  document.getElementById('emailInput').value = '';
  document.getElementById('preview').style.display = 'none';
  document.querySelector('.btn-download').disabled = true;
  convertedData = [];
}

let timeout;
document.getElementById('emailInput').addEventListener('input', function () {
  clearTimeout(timeout);
  timeout = setTimeout(() => {
    if (this.value.trim()) convertEmails();
  }, 1000);
});

// Gérer l'affichage de l'option "fichier unique" en fonction du nombre d'emails par feuille
document.getElementById('emailsPerSheet').addEventListener('change', function() {
  const value = parseInt(this.value, 10);
  const singleFileContainer = document.getElementById('singleFileContainer');
  const singleFileCheckbox = document.getElementById('singleFile');
  
  if (value === 0) {
    // Mode "tout en une seule feuille" : masquer l'option et forcer le mode fichier unique
    singleFileContainer.style.display = 'none';
    singleFileCheckbox.checked = true;
  } else {
    // Mode normal : afficher l'option
    singleFileContainer.style.display = 'block';
  }
});

function downloadOutlookCSV() {
  if (convertedData.length === 0) {
    alert("Aucune donnée à exporter.");
    return;
  }

  // Charger JSZip si ce n'est pas déjà fait
  if (typeof JSZip === 'undefined') {
    alert("La bibliothèque JSZip n'est pas chargée. Veuillez rafraîchir la page.");
    return;
  }

  const singleFile = document.getElementById('singleFile').checked;
  const emailsPerSheetValue = parseInt(document.getElementById('emailsPerSheet').value, 10);
  const totalEmails = convertedData.length;
  
  // Déterminer le mode d'export
  const isAllInOneSheet = emailsPerSheetValue === 0;
  if (isAllInOneSheet) {
    console.log("Mode une seule feuille sélectionné - tous les emails seront dans une seule feuille");
  }
  
  const emailsPerSheet = isAllInOneSheet ? totalEmails : (emailsPerSheetValue || 15);
  const expectedParts = isAllInOneSheet ? 1 : Math.ceil(totalEmails / emailsPerSheet);
  const now = new Date();
  const currentDate = now.toISOString().split("T")[0];
  const currentTime = now.getHours().toString().padStart(2, '0') + 
                   now.getMinutes().toString().padStart(2, '0') + 
                   now.getSeconds().toString().padStart(2, '0');
  
  console.log("=== Début de l'export (format Outlook) ===");
  console.log("Mode:", singleFile ? "Fichier unique" : "Fichiers séparés (ZIP)");
  console.log("Nombre d'emails par partie:", emailsPerSheet);
  console.log("Nombre total d'emails:", totalEmails);
  console.log("Nombre de parties attendues:", expectedParts);

  // Entête Outlook
  const headers = [
    "First Name","Middle Name","Last Name","Title","Suffix","Nickname","Given Yomi","Surname Yomi",
    "E-mail Address","E-mail 2 Address","E-mail 3 Address","Home Phone","Home Phone 2","Business Phone",
    "Business Phone 2","Mobile Phone","Car Phone","Other Phone","Primary"
  ];

  if (singleFile || isAllInOneSheet) {
    // Mode fichier unique : un seul CSV avec tous les contacts
    const separator = document.getElementById("useSemicolon").checked ? ";" : ",";
    console.log("Création d'un fichier CSV unique avec tous les contacts");

    const rows = convertedData.map(d => [
      d.prenom || "", "", d.nom || "", "", "", "", "", "",
      d.email || "", "", "", "", "", "", "", "", "", "", ""
    ]);

    const csvContent = [headers, ...rows]
      .map(row => row.map(value => `"${value}"`).join(separator))
      .join("\n");

    const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `contacts_outlook_${currentDate}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    console.log("Export terminé: 1 fichier CSV créé");
    window.alert("Fichier CSV unique créé avec tous les contacts");

  } else {
    // Mode fichiers séparés (ZIP)
    const zip = new JSZip();
    let fileCount = 0;
    const separator = document.getElementById("useSemicolon").checked ? ";" : ",";

    console.log("Création du fichier ZIP...");

    for (let i = 0; i < totalEmails; i += emailsPerSheet) {
      const currentEmails = convertedData.slice(i, i + emailsPerSheet);
      console.log(`Ajout fichier ${fileCount + 1} avec ${currentEmails.length} emails (position ${i} sur ${totalEmails})`);

      const rows = currentEmails.map(d => [
        d.prenom || "", "", d.nom || "", "", "", "", "", "",
        d.email || "", "", "", "", "", "", "", "", "", "", ""
      ]);

      const csvContent = [headers, ...rows]
        .map(row => row.map(value => `"${value}"`).join(separator))
        .join("\n");

      const fileName = `contacts_outlook_partie${(fileCount + 1).toString().padStart(2, '0')}.csv`;
      zip.file(fileName, "\uFEFF" + csvContent);
      fileCount++;
    }

    console.log("=== Création du ZIP terminée ===");
    console.log(`Nombre de fichiers ajoutés: ${fileCount}`);
    console.log(`Nombre de parties attendues: ${expectedParts}`);

    // Générer et télécharger le ZIP
    zip.generateAsync({type:"blob"})
      .then(function(content) {
        const url = URL.createObjectURL(content);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", `contacts_outlook_${currentDate}_${currentTime}.zip`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        window.alert(`Archive ZIP créée avec ${fileCount} fichier(s) CSV`);
      })
      .catch(function(error) {
        console.error("Erreur lors de la création du ZIP:", error);
        window.alert("Erreur lors de la création de l'archive ZIP");
      });
  }
}