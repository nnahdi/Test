// متغيرات عامة
let uploadedExcelFile = null;
let uploadedGoodSampleFile = null;
let uploadedNatsFile = null;

// تحميل القيم الافتراضية عند بدء التشغيل
document.addEventListener('DOMContentLoaded', function() {
    loadDefaultValues();
});

// تحميل الملفات الافتراضية من الخادم
async function loadDefaultFiles() {
    const statusSpan = document.getElementById('fileStatus');
    statusSpan.textContent = 'جاري تحميل الملفات الافتراضية...';
    
    try {
        // التحقق من وجود الملفات على الخادم
        const response = await fetch('/?handler=DefaultValues');
        if (response.ok) {
            const defaults = await response.json();
            
            // عرض القيم الافتراضية في الواجهة
            displayDefaultValues(defaults);
            statusSpan.textContent = '✓ تم تحميل الملفات الافتراضية بنجاح';
            statusSpan.className = 'align-self-center text-success';
        }
    } catch (error) {
        console.error('Error loading default files:', error);
        statusSpan.textContent = '✗ خطأ في تحميل الملفات الافتراضية';
        statusSpan.className = 'align-self-center text-danger';
    }
}

// تحميل القيم الافتراضية من ملف النموذج
async function loadDefaultValues() {
    try {
        const response = await fetch('/?handler=DefaultValues');
        if (response.ok) {
            const defaults = await response.json();
            displayDefaultValues(defaults);
        }
    } catch (error) {
        console.error('Error loading default values:', error);
    }
}

// عرض القيم الافتراضية في الواجهة
function displayDefaultValues(defaults) {
    const container = document.getElementById('defaultValuesContainer');
    container.innerHTML = '';
    
    if (!defaults || Object.keys(defaults).length === 0) {
        container.innerHTML = '<div class="col-12 text-muted">لا توجد قيم افتراضية متاحة</div>';
        return;
    }
    
    let colIndex = 0;
    for (const [key, value] of Object.entries(defaults)) {
        const colClass = colIndex % 3 === 0 ? 'col-md-4' : 'col-md-4';
        
        const div = document.createElement('div');
        div.className = colClass;
        div.innerHTML = `
            <label for="default_${key}" class="form-label">${key}</label>
            <input type="text" class="form-control default-value-input" 
                   id="default_${key}" 
                   data-column="${key}" 
                   value="${value || ''}"
                   placeholder="القيمة الافتراضية">
        `;
        
        container.appendChild(div);
        colIndex++;
    }
}

// معالجة رفع الملف وعرض المعاينة
document.getElementById('excelFile').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    uploadedExcelFile = file;
    
    const formData = new FormData();
    formData.append('excelFile', file);
    
    try {
        const response = await fetch('/?handler=Preview', {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const data = await response.json();
            showPreview(data.columns, data.rows);
            document.getElementById('fileStatus').textContent = `✓ تم رفع الملف: ${data.rowCount} صف، ${data.columns.length} عمود`;
            document.getElementById('fileStatus').className = 'align-self-center text-success';
        }
    } catch (error) {
        console.error('Error previewing file:', error);
        document.getElementById('fileStatus').textContent = '✗ خطأ في معاينة الملف';
        document.getElementById('fileStatus').className = 'align-self-center text-danger';
    }
});

// معالجة رفع ملف النموذج الصحيح
document.getElementById('goodSampleFile').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    uploadedGoodSampleFile = file;
    document.getElementById('fileStatus').textContent = '✓ تم رفع ملف النموذج الصحيح';
    document.getElementById('fileStatus').className = 'align-self-center text-success';
});

// معالجة رفع ملف الجنسيات
document.getElementById('natsFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    uploadedNatsFile = file;
    document.getElementById('fileStatus').textContent = '✓ تم رفع ملف الجنسيات';
    document.getElementById('fileStatus').className = 'align-self-center text-success';
});

// عرض معاينة البيانات
function showPreview(columns, rows) {
    const previewDiv = document.getElementById('previewDiv');
    const table = document.getElementById('previewTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    
    // مسح المحتوى السابق
    thead.innerHTML = '';
    tbody.innerHTML = '';
    
    // إنشاء رؤوس الأعمدة
    const headerRow = document.createElement('tr');
    columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        th.className = 'text-nowrap';
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    
    // إنشاء صفوف البيانات
    rows.forEach(row => {
        const tr = document.createElement('tr');
        columns.forEach(col => {
            const td = document.createElement('td');
            td.textContent = row[col] || '';
            td.className = 'text-nowrap';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    
    previewDiv.style.display = 'block';
}

// معالجة الملف الرئيسي
async function processFile() {
    if (!uploadedExcelFile) {
        showError('يرجى اختيار ملف Excel للمعالجة');
        return;
    }
    
    // إظهار شاشة التحميل
    document.getElementById('loadingDiv').style.display = 'block';
    document.getElementById('resultsDiv').style.display = 'none';
    document.getElementById('downloadDiv').style.display = 'none';
    document.getElementById('errorDiv').style.display = 'none';
    
    // جمع الإعدادات
    const bookingStart = parseInt(document.getElementById('bookingStart').value) || 1;
    const identityStart = parseInt(document.getElementById('identityStart').value) || 1000000000;
    const defaultPhone = document.getElementById('defaultPhone').value || '966500000000';
    
    // جمع القيم الافتراضية المخصصة
    const defaultValues = {};
    document.querySelectorAll('.default-value-input').forEach(input => {
        if (input.value) {
            defaultValues[input.dataset.column] = input.value;
        }
    });
    
    const formData = new FormData();
    formData.append('excelFile', uploadedExcelFile);
    formData.append('bookingStart', bookingStart);
    formData.append('identityStart', identityStart);
    formData.append('defaultPhone', defaultPhone);
    
    if (uploadedGoodSampleFile) {
        formData.append('goodSampleFile', uploadedGoodSampleFile);
    }
    
    if (uploadedNatsFile) {
        formData.append('natsFile', uploadedNatsFile);
    }
    
    try {
        const response = await fetch('/?handler=Process', {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const result = await response.json();
            displayResults(result);
        } else {
            const errorData = await response.json();
            showError(errorData.error || errorData.errors?.join(', ') || 'حدث خطأ غير معروف');
        }
    } catch (error) {
        console.error('Error processing file:', error);
        showError('حدث خطأ أثناء معالجة الملف');
    } finally {
        document.getElementById('loadingDiv').style.display = 'none';
    }
}

// عرض نتائج المعالجة
function displayResults(result) {
    document.getElementById('totalRows').textContent = result.totalRows;
    document.getElementById('totalColumns').textContent = result.totalColumns;
    document.getElementById('correctedValues').textContent = result.correctedValues;
    document.getElementById('invalidPhones').textContent = result.invalidPhoneNumbers;
    document.getElementById('matchedNats').textContent = result.matchedNationalities;
    document.getElementById('addedColumns').textContent = result.addedColumns;
    document.getElementById('extraColumns').textContent = result.extraColumnsKept;
    
    document.getElementById('resultsDiv').style.display = 'block';
    
    // إعداد زر التحميل
    const downloadBtn = document.getElementById('downloadBtn');
    downloadBtn.href = '/?handler=Download';
    downloadBtn.download = result.fileName;
    document.getElementById('downloadDiv').style.display = 'block';
}

// عرض رسالة خطأ
function showError(message) {
    const errorDiv = document.getElementById('errorDiv');
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
}
