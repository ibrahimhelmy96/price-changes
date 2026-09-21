<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>طباعة الباركود - Epson TM-T20III</title>
    <!-- مكتبات معالجة الإكسيل والباركود والـ PDF -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.5/dist/JsBarcode.all.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>

    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f4f7f6;
            margin: 0;
            padding: 20px;
            display: flex;
            justify-content: center;
        }
        .container {
            background-color: #fff;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
            width: 100%;
            max-width: 500px;
        }
        h2 {
            text-align: center;
            color: #333;
            margin-bottom: 20px;
        }
        .file-group {
            margin-bottom: 15px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #555;
        }
        input[type="file"] {
            width: 100%;
            padding: 8px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
        }
        button {
            width: 100%;
            padding: 12px;
            background-color: #28a745;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            margin-top: 10px;
        }
        button:hover {
            background-color: #218838;
        }
        #status {
            margin-top: 15px;
            text-align: center;
            font-weight: bold;
        }
    </style>
</head>
<body>

<div class="container">
    <h2>طابعة الباركود (Epson)</h2>

    <div class="file-group">
        <label>شيت A (المخزون الأساسي):</label>
        <input type="file" id="fileA" accept=".xlsx, .xls">
    </div>

    <div class="file-group">
        <label>شيت B (الأسعار المعدلة):</label>
        <input type="file" id="fileB" accept=".xlsx, .xls">
    </div>

    <button onclick="processFiles()">مطابقة وإنشاء ملف الباركود</button>

    <div id="status"></div>
</div>

<svg id="barcodeCanvas" style="display:none;"></svg>

<script>
    const { jsPDF } = window.jspdf;

    function readExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet);
                resolve(json);
            };
            reader.onerror = (error) => reject(error);
            reader.readAsArrayBuffer(file);
        });
    }

    // دالة للحصول على قيمة الخلية بغض النظر عن حالة الأحرف الكبيرة أو الصغيرة أو اللغة
    function getValue(row, possibleKeys) {
        for (let key in row) {
            const cleanKey = key.trim().toLowerCase();
            for (let p of possibleKeys) {
                if (cleanKey === p.toLowerCase()) {
                    return row[key];
                }
            }
        }
        return null;
    }

    async function processFiles() {
        const fileAInput = document.getElementById('fileA').files[0];
        const fileBInput = document.getElementById('fileB').files[0];
        const status = document.getElementById('status');

        if (!fileAInput || !fileBInput) {
            alert('برجاء اختيار الملفين أولاً');
            return;
        }

        status.style.color = 'blue';
        status.innerText = 'جاري المعالجة والمطابقة...';

        try {
            const dataA = await readExcel(fileAInput); // شيت المخزون (A)
            const dataB = await readExcel(fileBInput); // شيت الأسعار المعدلة (B)

            // 1. قراءة كود الصنف واسمه من شيت A
            const mapA = new Map();
            dataA.forEach(row => {
                const ref = getValue(row, ['Product Ref', 'ProductRef', 'ref', 'الرمز', 'الكود']);
                const name = getValue(row, ['Product Name', 'ProductName', 'الاسم', 'اسم الصنف']);
                if (ref !== null && ref !== undefined) {
                    mapA.set(String(ref).trim(), name || '');
                }
            });

            // 2. مطابقة الأكواد من شيت B مع شيت A وجلب السعر من B والاسم من A
            const matchedData = [];
            dataB.forEach(row => {
                const ref = getValue(row, ['product ref', 'Product Ref', 'ref', 'الكود']);
                const price = getValue(row, ['price', 'Price', 'السعر']);

                if (ref !== null && ref !== undefined) {
                    const cleanRef = String(ref).trim();
                    if (mapA.has(cleanRef)) {
                        matchedData.push({
                            ref: cleanRef,
                            price: price !== null ? parseFloat(price).toFixed(2) : '',
                            name: mapA.get(cleanRef)
                        });
                    }
                }
            });

            if (matchedData.length === 0) {
                status.style.color = 'red';
                status.innerText = 'لم يتم العثور على أكواد مطابقة بين الشيتين!';
                return;
            }

            generatePDF(matchedData);
            status.style.color = 'green';
            status.innerText = `تم مطابقة ${matchedData.length} صنف وإنشاء ملف الطباعة بنجاح!`;

        } catch (error) {
            console.error(error);
            status.style.color = 'red';
            status.innerText = 'حدث خطأ أثناء معالجة الملفات.';
        }
    }

    function generatePDF(items) {
        const labelWidth = 80;  // عرض رول الطابعة مم (Epson TM-T20III)
        const labelHeight = 45; // ارتفاع ملصق الباركود
        const totalHeight = items.length * labelHeight;

        const doc = new jsPDF({
            orientation: 'p',
            unit: 'mm',
            format: [labelWidth, totalHeight]
        });

        const svg = document.getElementById('barcodeCanvas');

        items.forEach((item, index) => {
            const topY = index * labelHeight;

            // 1. اسم الصنف من شيت A (بخط واضح)
            doc.setFontSize(10);
            doc.setFont("helvetica", "bold");
            doc.text(String(item.name).substring(0, 35), labelWidth / 2, topY + 8, { align: 'center' });

            // 2. الباركود كود 128 برقم Product Ref
            JsBarcode(svg, item.ref, {
                format: "CODE128",
                displayValue: false,
                margin: 0
            });
            
            const xml = new XMLSerializer().serializeToString(svg);
            const svg64 = btoa(unescape(encodeURIComponent(xml)));
            const image64 = 'data:image/svg+xml;base64,' + svg64;

            doc.addImage(image64, 'SVG', 5, topY + 12, 70, 18);

            // 3. Product Ref على اليسار والسعر من شيت B على اليمين
            doc.setFontSize(8);
            doc.setFont("helvetica", "normal");
            
            doc.text(`Ref: ${item.ref}`, 5, topY + 35, { align: 'left' });
            doc.text(`Price: ${item.price}`, labelWidth - 5, topY + 35, { align: 'right' });

            // خط فاصل بين الملصقات
            if (index < items.length - 1) {
                doc.setLineDashPattern([1, 1], 0);
                doc.line(2, topY + labelHeight, labelWidth - 2, topY + labelHeight);
            }
        });

        // فتح ملف الـ PDF الجاهز للطباعة مباشرة
        doc.output('dataurlnewwindow');
    }
</script>

</body>
</html>
