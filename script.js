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
  const ldifBtn = document.querySelector('.btn-ldif');

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
  ldifBtn.disabled = false;
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

function btoa_utf8(str) {
  return btoa(unescape(encodeURIComponent(str)));
}

function downloadLDIF() {
  if (convertedData.length === 0) {
    alert("Aucune donnée à exporter.");
    return;
  }

  let ldifContent = '';
  
  convertedData.forEach(data => {
    const displayName = `${data.prenom} ${data.nom}`;
    const dn = `cn=${data.prenom} ${data.nom},mail=${data.email}`;
    
    ldifContent += `dn:: ${btoa_utf8(dn)}\n`;
    ldifContent += 'objectClass: top\n';
    ldifContent += 'objectClass: inetOrgPerson\n';
    ldifContent += 'objectClass: mozillaAbPersonAlpha\n';
    ldifContent += `displayname:: ${btoa_utf8(displayName)}\n`;
    ldifContent += `mail: ${data.email}\n`;
    ldifContent += `sn: ${data.nom}\n`;
    ldifContent += `givenname: ${data.prenom}\n`;
    ldifContent += '\n';
  });

  const currentDate = new Date().toISOString().split("T")[0];
  downloadFile(ldifContent, `contacts_${currentDate}.ldif`, 'application/x-ldif;charset=utf-8');
}

function downloadFile(content, filename, type) {
  const blob = new Blob(["\uFEFF" + content], { type: type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}