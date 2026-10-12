// Míchání studentských dvojic - aplikační logika
(function () {
  'use strict';

  // Stav aplikace
  let allStudents = [];
  let selectedGroup = null;
  let currentMixedGroups = []; // pole polí studentů, např. [[s1, s2], [s3, s4, s5]]
  let currentSourceName = '';

  // DOM prvky
  const statusDot = document.getElementById('statusDot');
  const statusText = document.getElementById('statusText');
  const btnReloadDefault = document.getElementById('btnReloadDefault');
  const fileInput = document.getElementById('fileInput');
  const dropZone = document.getElementById('dropZone');
  
  const groupSelect = document.getElementById('groupSelect');
  const btnMix = document.getElementById('btnMix');
  const groupInfo = document.getElementById('groupInfo');
  const pillStudentCount = document.getElementById('pillStudentCount');
  const pillPairBreakdown = document.getElementById('pillPairBreakdown');
  const pillOddNotice = document.getElementById('pillOddNotice');
  const btnToggleStudentList = document.getElementById('btnToggleStudentList');
  const studentRoster = document.getElementById('studentRoster');
  const rosterList = document.getElementById('rosterList');

  const resultsSection = document.getElementById('resultsSection');
  const resultsSubtitle = document.getElementById('resultsSubtitle');
  const pairsGrid = document.getElementById('pairsGrid');
  const btnRemix = document.getElementById('btnRemix');
  const btnCopyText = document.getElementById('btnCopyText');
  const btnDownloadTxt = document.getElementById('btnDownloadTxt');
  const previewLineCount = document.getElementById('previewLineCount');
  const txtPreviewBox = document.getElementById('txtPreviewBox');
  const toast = document.getElementById('toast');

  // Inicializace
  document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    loadDefaultData();
  });

  // Registrace posluchačů událostí
  function initEventListeners() {
    btnReloadDefault.addEventListener('click', loadDefaultData);

    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) handleUploadedFile(file);
    });

    // Drag & Drop
    ['dragenter', 'dragover'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
      });
    });

    dropZone.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        handleUploadedFile(files[0]);
      }
    });

    dropZone.addEventListener('click', () => {
      fileInput.click();
    });

    groupSelect.addEventListener('change', onGroupSelectChange);
    btnMix.addEventListener('click', mixStudents);
    btnRemix.addEventListener('click', mixStudents);

    btnToggleStudentList.addEventListener('click', () => {
      const isHidden = studentRoster.style.display === 'none';
      studentRoster.style.display = isHidden ? 'block' : 'none';
      btnToggleStudentList.textContent = isHidden ? 'Skrýt seznam studentů' : `Zobrazit seznam studentů (${getRosterCount()})`;
    });

    btnCopyText.addEventListener('click', copyResultsToClipboard);
    btnDownloadTxt.addEventListener('click', downloadResultsAsTxt);
  }

  // Načtení výchozích dat (nejprve fetch souboru VND26_studenti.xlsx, případně fallback na window.DEFAULT_STUDENTS)
  async function loadDefaultData() {
    updateStatus('loading', 'Načítám výchozí soubor VND26_studenti.xlsx...');
    
    try {
      // Zkusíme stáhnout soubor přes fetch (standardní cesta pro GitHub Pages i lokální web server)
      const response = await fetch('./VND26_studenti.xlsx', { cache: 'no-cache' });
      if (response.ok) {
        const arrayBuffer = await response.arrayBuffer();
        parseExcelBuffer(arrayBuffer, 'VND26_studenti.xlsx');
        return;
      }
    } catch (err) {
      console.warn('Fetch VND26_studenti.xlsx se nezdařil (časté při lokálním otevření přes file://):', err);
    }

    // Fallback na vestavěná data z students_data.js
    if (window.DEFAULT_STUDENTS && Array.isArray(window.DEFAULT_STUDENTS) && window.DEFAULT_STUDENTS.length > 0) {
      setStudentsData(window.DEFAULT_STUDENTS, 'VND26_studenti.xlsx (vestavěná data)');
    } else {
      updateStatus('error', 'Nepodařilo se automaticky načíst VND26_studenti.xlsx. Nahrajte prosím soubor tlačítkem výše.');
    }
  }

  // Zpracování nahraného souboru
  function handleUploadedFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      showToast('Vyberte prosím soubor ve formátu Excel (.xlsx nebo .xls)', true);
      return;
    }

    updateStatus('loading', `Načítám soubor ${file.name}...`);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        parseExcelBuffer(e.target.result, file.name);
      } catch (err) {
        console.error('Chyba při čtení Excelu:', err);
        updateStatus('error', `Chyba při zpracování souboru: ${err.message}`);
        showToast('Chyba při zpracování souboru Excel', true);
      }
    };
    reader.onerror = () => {
      updateStatus('error', 'Chyba při načítání souboru z disku.');
    };
    reader.readAsArrayBuffer(file);
  }

  // Parsování ArrayBufferu pomocí SheetJS
  function parseExcelBuffer(arrayBuffer, sourceName) {
    if (typeof XLSX === 'undefined') {
      throw new Error('Knihovna XLSX (SheetJS) není načtena.');
    }

    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Převedeme na pole polí (řádků)
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    if (!rows || rows.length < 2) {
      throw new Error('Soubor neobsahuje žádné datové řádky.');
    }

    // Zjištění, zda 1. řádek obsahuje hlavičku
    let startIndex = 0;
    const firstRowStr = (rows[0] || []).join(' ').toLowerCase();
    if (firstRowStr.includes('jmén') || firstRowStr.includes('name') || firstRowStr.includes('skupin')) {
      startIndex = 1;
    }

    const parsedStudents = [];

    for (let r = startIndex; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length < 2) continue;

      const col1 = String(row[0] || '').trim();
      const col2 = String(row[1] || '').trim();
      const col3 = row[2] !== undefined ? String(row[2]).trim() : '';

      if (!col1 && !col2) continue; // prázdný řádek

      // Zpracování skupiny:
      // Může být přímo číslo "42" nebo text "B261-11VND studijní skupina 42"
      let groupVal = col3;
      const numMatch = col3.match(/(\d+)$/);
      if (numMatch) {
        groupVal = parseInt(numMatch[1], 10);
      }

      parsedStudents.push({
        firstName: col1,
        lastName: col2,
        group: groupVal
      });
    }

    if (parsedStudents.length === 0) {
      throw new Error('Ze souboru se nepodařilo načíst žádné studenty.');
    }

    setStudentsData(parsedStudents, sourceName);
  }

  // Uložení dat a aktualizace UI
  function setStudentsData(students, sourceName) {
    allStudents = students;
    currentSourceName = sourceName;

    // Zjistíme unikátní skupiny
    const groupsSet = new Set();
    allStudents.forEach(s => {
      if (s.group !== '' && s.group !== null && s.group !== undefined) {
        groupsSet.add(s.group);
      }
    });

    const sortedGroups = Array.from(groupsSet).sort((a, b) => {
      const numA = Number(a);
      const numB = Number(b);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return String(a).localeCompare(String(b), 'cs');
    });

    updateStatus('ready', `Aktivní zdroj: ${sourceName} (${allStudents.length} studentů v ${sortedGroups.length} skupinách)`);
    populateGroupSelect(sortedGroups);
    showToast(`Úspěšně načteno ${allStudents.length} studentů`);
  }

  // Naplnění výběrové roletky
  function populateGroupSelect(groups) {
    groupSelect.innerHTML = '<option value="">-- Vyberte třídu / studijní skupinu --</option>';

    groups.forEach(g => {
      const count = allStudents.filter(s => s.group === g).length;
      const opt = document.createElement('option');
      opt.value = g;
      opt.textContent = `Skupina ${g} (${count} studentů)`;
      groupSelect.appendChild(opt);
    });

    groupSelect.disabled = false;

    // Reset výběru a výsledků
    selectedGroup = null;
    groupInfo.style.display = 'none';
    studentRoster.style.display = 'none';
    btnMix.disabled = true;
    resultsSection.style.display = 'none';
  }

  // Změna vybrané skupiny
  function onGroupSelectChange() {
    const val = groupSelect.value;
    if (!val) {
      selectedGroup = null;
      btnMix.disabled = true;
      groupInfo.style.display = 'none';
      studentRoster.style.display = 'none';
      resultsSection.style.display = 'none';
      return;
    }

    // Skupina může být uložena jako číslo nebo string
    const matchNum = Number(val);
    selectedGroup = isNaN(matchNum) ? val : matchNum;

    const groupStudents = getStudentsForSelectedGroup();
    const count = groupStudents.length;

    if (count < 2) {
      showToast('Vybraná skupina má méně než 2 studenty.', true);
      btnMix.disabled = true;
      groupInfo.style.display = 'none';
      return;
    }

    btnMix.disabled = false;
    groupInfo.style.display = 'flex';

    // Výpočet rozpisu
    const isOdd = count % 2 !== 0;
    const pairCount = isOdd ? (count - 3) / 2 : count / 2;

    pillStudentCount.textContent = `Studentů: ${count}`;
    if (isOdd) {
      pillPairBreakdown.textContent = `Rozpis: ${pairCount} dvojic + 1 trojice`;
      pillOddNotice.style.display = 'inline-block';
    } else {
      pillPairBreakdown.textContent = `Rozpis: ${pairCount} dvojic`;
      pillOddNotice.style.display = 'none';
    }

    btnToggleStudentList.textContent = `Zobrazit seznam studentů (${count})`;
    renderRoster(groupStudents);
    resultsSection.style.display = 'none';
  }

  function getStudentsForSelectedGroup() {
    if (selectedGroup === null) return [];
    return allStudents.filter(s => s.group === selectedGroup);
  }

  function getRosterCount() {
    return getStudentsForSelectedGroup().length;
  }

  function renderRoster(students) {
    rosterList.innerHTML = '';
    // Setřídíme abecedně podle příjmení pro přehlednost
    const sorted = [...students].sort((a, b) => a.lastName.localeCompare(b.lastName, 'cs'));
    sorted.forEach(s => {
      const li = document.createElement('li');
      li.textContent = `${s.lastName} ${s.firstName}`;
      rosterList.appendChild(li);
    });
  }

  // Náhodné promíchání (Fisher-Yates) a sestavení dvojic / trojice
  function mixStudents() {
    const students = getStudentsForSelectedGroup();
    const n = students.length;

    if (n < 2) {
      showToast('Pro vytvoření dvojic je potřeba alespoň 2 studenti.', true);
      return;
    }

    // Zamíchání kopií pole
    const shuffled = [...students];
    for (let i = shuffled.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }

    const groups = [];
    if (n === 3) {
      groups.push([shuffled[0], shuffled[1], shuffled[2]]);
    } else if (n % 2 === 0) {
      // Sudý počet: pouze dvojice
      for (let i = 0; i < n; i += 2) {
        groups.push([shuffled[i], shuffled[i + 1]]);
      }
    } else {
      // Lichý počet: dvojice a na konci 1 trojice
      for (let i = 0; i < n - 3; i += 2) {
        groups.push([shuffled[i], shuffled[i + 1]]);
      }
      groups.push([shuffled[n - 3], shuffled[n - 2], shuffled[n - 1]]);
    }

    currentMixedGroups = groups;
    renderResults(groups, n);
  }

  // Zobrazení výsledků
  function renderResults(groups, totalStudents) {
    pairsGrid.innerHTML = '';
    const isOdd = totalStudents % 2 !== 0;
    const pairCount = isOdd ? groups.length - 1 : groups.length;

    resultsSubtitle.textContent = isOdd 
      ? `Skupina ${selectedGroup}: celkem ${totalStudents} studentů rozděleno do ${pairCount} dvojic a 1 trojice`
      : `Skupina ${selectedGroup}: celkem ${totalStudents} studentů rozděleno do ${pairCount} dvojic`;

    groups.forEach((group, idx) => {
      const isTriplet = group.length === 3;
      const card = document.createElement('div');
      card.className = `pair-card ${isTriplet ? 'triplet' : ''}`;

      const header = document.createElement('div');
      header.className = 'pair-card-header';
      
      const badge = document.createElement('span');
      badge.className = 'pair-badge';
      badge.textContent = isTriplet ? '⭐ Trojice (lichý počet)' : `Dvojice #${idx + 1}`;
      header.appendChild(badge);
      card.appendChild(header);

      group.forEach(student => {
        const item = document.createElement('div');
        item.className = 'student-list-item';

        const avatar = document.createElement('div');
        avatar.className = 'student-avatar';
        // Iniciály
        const ini1 = (student.firstName || '').charAt(0).toUpperCase();
        const ini2 = (student.lastName || '').charAt(0).toUpperCase();
        avatar.textContent = `${ini1}${ini2}`;

        const nameSpan = document.createElement('span');
        nameSpan.textContent = `${student.firstName} ${student.lastName}`;

        item.appendChild(avatar);
        item.appendChild(nameSpan);
        card.appendChild(item);
      });

      pairsGrid.appendChild(card);
    });

    // Aktualizace náhledu TXT souboru
    updateTxtPreview(groups);

    // Zobrazit sekci výsledků a hladce odrolovat
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    showToast('Studenti byli náhodně namícháni!');
  }

  // Vygenerování textu pro TXT: jeden řádek = jedna dvojice/trojice
  function generateTxtContent(groups) {
    return groups.map(group => {
      return group.map(s => `${s.firstName} ${s.lastName}`).join(', ');
    }).join('\r\n');
  }

  function updateTxtPreview(groups) {
    const txt = generateTxtContent(groups);
    txtPreviewBox.textContent = txt;
    previewLineCount.textContent = `${groups.length} ${groups.length === 1 ? 'řádek' : (groups.length < 5 ? 'řádky' : 'řádků')}`;
  }

  // Uložení TXT do počítače
  function downloadResultsAsTxt() {
    if (!currentMixedGroups || currentMixedGroups.length === 0) {
      showToast('Nejprve zamíchejte studenty.', true);
      return;
    }

    const textContent = generateTxtContent(currentMixedGroups);
    // Přidáme UTF-8 BOM (\uFEFF) pro bezproblémové zobrazení českých znaků v Poznámkovém bloku i Excelu
    const blob = new Blob(['\uFEFF' + textContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = `skupina_${selectedGroup}_dvojice.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast(`Soubor skupina_${selectedGroup}_dvojice.txt byl stažen`);
  }

  // Kopírování do schránky
  async function copyResultsToClipboard() {
    if (!currentMixedGroups || currentMixedGroups.length === 0) {
      showToast('Nejprve zamíchejte studenty.', true);
      return;
    }

    const textContent = generateTxtContent(currentMixedGroups);

    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(textContent);
      } else {
        // Fallback pro starší prohlížeče
        const textarea = document.createElement('textarea');
        textarea.value = textContent;
        textarea.style.position = 'fixed';
        textarea.style.opacity = '0';
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
      }
      showToast('Výsledky zkopírovány do schránky');
    } catch (err) {
      console.error('Chyba při kopírování:', err);
      showToast('Kopírování do schránky se nezdařilo.', true);
    }
  }

  // Pomocné funkce
  function updateStatus(state, message) {
    statusDot.className = `status-indicator ${state}`;
    statusText.textContent = message;
  }

  let toastTimeout = null;
  function showToast(message, isError = false) {
    if (toastTimeout) clearTimeout(toastTimeout);
    toast.textContent = message;
    toast.style.backgroundColor = isError ? '#ef4444' : '#1e293b';
    toast.classList.add('show');
    toastTimeout = setTimeout(() => {
      toast.classList.remove('show');
    }, 3200);
  }

})();
