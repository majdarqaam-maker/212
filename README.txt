<script type="text/javascript">
        var gk_isXlsx = false;
        var gk_xlsxFileLookup = {};
        var gk_fileData = {};
        function filledCell(cell) {
          return cell !== '' && cell != null;
        }
        function loadFileData(filename) {
        if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
            try {
                var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to JSON to filter blank rows
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
                // Filter out blank rows (rows where all cells are empty, null, or undefined)
                var filteredData = jsonData.filter(row => row.some(filledCell));

                // Heuristic to find the header row by ignoring rows with fewer filled cells than the next row
                var headerRowIndex = filteredData.findIndex((row, index) =>
                  row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
                );
                // Fallback
                if (headerRowIndex === -1 || headerRowIndex > 25) {
                  headerRowIndex = 0;
                }

                // Convert filtered JSON back to CSV
                var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex)); // Create a new sheet from filtered array of arrays
                csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
                return csv;
            } catch (e) {
                console.error(e);
                return "";
            }
        }
        return gk_fileData[filename] || "";
        }
        </script><!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تقرير مركز الاتصال</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Amiri&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Amiri', sans-serif;
            background: linear-gradient(135deg, #e0f7fa, #b2ebf2);
            display: flex;
            justify-content: center;
            align-items: flex-start;
            min-height: 100vh;
            margin: 0;
            padding: 20px;
            color: #333;
            direction: rtl;
        }
        .form-container {
            background: rgba(255, 255, 255, 0.95);
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.1);
            width: 100%;
            max-width: 600px;
        }
        h2 {
            text-align: center;
            color: #00bcd4;
            font-size: 1.8em;
            margin-bottom: 20px;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            color: #555;
            font-weight: 500;
        }
        input, select, textarea {
            width: 100%;
            padding: 12px;
            border: 1px solid #00bcd4;
            border-radius: 8px;
            box-sizing: border-box;
            background: #f0faff;
            font-family: 'Amiri', sans-serif;
            text-align: right;
        }
        input:focus, select:focus, textarea:focus {
            border-color: #0097a7;
            box-shadow: 0 0 8px rgba(0, 188, 212, 0.3);
            outline: none;
        }
        textarea {
            height: 120px;
            resize: vertical;
        }
        button {
            width: 100%;
            padding: 12px;
            background: linear-gradient(90deg, #00bcd4, #0097a7);
            color: white;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        button:hover {
            background: linear-gradient(90deg, #0097a7, #00bcd4);
        }
        #displayDate, #displayStartDate, #displayEndDate {
            margin-top: 5px;
            color: #00bcd4;
            font-weight: bold;
        }
        .hidden {
            display: none;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }
        th, td {
            padding: 10px;
            border: 1px solid #00bcd4;
            text-align: right;
        }
        th {
            background: #e0f7fa;
            color: #00796b;
        }
        .time-range {
            display: flex;
            gap: 10px;
        }
        .time-range input {
            width: 48%;
        }
        .chart-container {
            margin: 20px 0;
            padding: 10px;
            background: #f0faff;
            border-radius: 8px;
            border: 1px solid #00bcd4;
            text-align: center;
        }
        canvas {
            max-width: 100%;
        }
        .notes-section h3 {
            color: #00796b;
            margin-bottom: 10px;
        }
    </style>
</head>
<body>
    <div class="form-container">
        <h2>تقرير مركز الاتصال</h2>
        <form id="reportForm">
            <table>
                <tr><th colspan="2">المعلومات الأساسية</th></tr>
                <tr>
                    <td>نوع التقرير</td>
                    <td>
                        <select id="reportType" name="reportType" required>
                            <option value="Daily">يومي</option>
                            <option value="Weekly">أسبوعي</option>
                            <option value="Monthly">شهري</option>
                        </select>
                    </td>
                </tr>
            </table>

            <table id="singleDateGroup">
                <tr><th colspan="2">التاريخ</th></tr>
                <tr>
                    <td>اختر التاريخ</td>
                    <td>
                        <input type="date" id="reportDate" name="reportDate">
                        <div id="displayDate"></div>
                    </td>
                </tr>
            </table>

            <table id="dateRangeGroup" class="hidden">
                <tr><th colspan="2">نطاق التاريخ</th></tr>
                <tr>
                    <td>تاريخ البداية</td>
                    <td>
                        <input type="date" id="startDate" name="startDate">
                        <div id="displayStartDate"></div>
                    </td>
                </tr>
                <tr>
                    <td>تاريخ النهاية</td>
                    <td>
                        <input type="date" id="endDate" name="endDate">
                        <div id="displayEndDate"></div>
                    </td>
                </tr>
            </table>

            <table>
                <tr><th colspan="2">المشروع والقسم</th></tr>
                <tr>
                    <td>المشروع</td>
                    <td>
                        <select id="project" name="project" required>
                            <option value="">اختر المشروع</option>
                            <option value="TMC">TMC</option>
                            <option value="NUZUL">نزل</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td>القسم</td>
                    <td><input type="text" id="department" name="department" value="قسم مركز الاتصال - TMC"></td>
                </tr>
            </table>

            <table id="queuesContainer" class="hidden">
                <tr><th colspan="2">الطوابير</th></tr>
                <tr><td colspan="2"><div id="queuesList"></div></td></tr>
            </table>

            <table>
                <tr><th colspan="2">إحصائيات المكالمات</th></tr>
                <tr>
                    <td>المكالمات المستلمة</td>
                    <td><input type="number" id="receivedCalls" name="receivedCalls" min="0"></td>
                </tr>
                <tr>
                    <td>المكالمات المجابة</td>
                    <td><input type="number" id="answeredCalls" name="answeredCalls" min="0"></td>
                </tr>
                <tr>
                    <td>المكالمات غير المجابة</td>
                    <td><input type="number" id="unansweredCalls" name="unansweredCalls" min="0"></td>
                </tr>
                <tr>
                    <td>المكالمات المهجورة</td>
                    <td><input type="number" id="abandonedCalls" name="abandonedCalls" min="0"></td>
                </tr>
                <tr>
                    <td>المكالمات الصادرة</td>
                    <td><input type="number" id="outgoingCalls" name="outgoingCalls" min="0"></td>
                </tr>
            </table>

            <div class="chart-container">
                <h3>إحصائيات المكالمات</h3>
                <canvas id="callsChart"></canvas>
            </div>

            <table id="extraFields" class="hidden">
                <tr><th colspan="2">مؤشرات الأداء</th></tr>
                <tr>
                    <td>متوسط مدة المكالمة</td>
                    <td><input type="text" id="avgCallDuration" name="avgCallDuration"></td>
                </tr>
                <tr>
                    <td>متوسط وقت الانتظار</td>
                    <td><input type="text" id="avgWaitTime" name="avgWaitTime"></td>
                </tr>
                <tr>
                    <td>أوقات الذروة</td>
                    <td>
                        <div class="time-range">
                            <input type="time" id="peakStart" name="peakStart" placeholder="من">
                            <input type="time" id="peakEnd" name="peakEnd" placeholder="إلى">
                        </div>
                    </td>
                </tr>
                <tr id="tmcExtraCleaning" class="hidden">
                    <td>التنظيف</td>
                    <td><input type="text" id="cleaning" name="cleaning"></td>
                </tr>
                <tr id="tmcExtraMaintenance" class="hidden">
                    <td>الصيانة</td>
                    <td><input type="text" id="maintenance" name="maintenance"></td>
                </tr>
                <tr id="tmcExtraPropertyOwners" class="hidden">
                    <td>أصحاب العقارات</td>
                    <td><input type="text" id="propertyOwners" name="propertyOwners"></td>
                </tr>
                <tr id="tmcExtraBooking" class="hidden">
                    <td>الحجز</td>
                    <td><input type="text" id="booking" name="booking"></td>
                </tr>
                <tr id="tmcExtraBookingClassification" class="hidden">
                    <td>تصنيف الحجز</td>
                    <td><input type="text" id="bookingClassification" name="bookingClassification"></td>
                </tr>
                <tr id="tmcExtraMonthlyBooking" class="hidden">
                    <td>الحجوزات الشهرية</td>
                    <td><input type="text" id="monthlyBooking" name="monthlyBooking"></td>
                </tr>
                <tr id="tmcExtraDailyBooking" class="hidden">
                    <td>الحجوزات اليومية</td>
                    <td><input type="text" id="dailyBooking" name="dailyBooking"></td>
                </tr>
            </table>

            <table id="ticketsTable">
                <tr><th colspan="2">التذاكر</th></tr>
                <tr id="tmcComplaints" class="hidden">
                    <td>الشكاوى</td>
                    <td><input type="number" id="complaints" name="complaints" min="0"></td>
                </tr>
                <tr id="tmcContactRequests" class="hidden">
                    <td>طلبات التواصل</td>
                    <td><input type="number" id="contactRequestsInput" name="contactRequestsInput" min="0"></td>
                </tr>
                <tr id="nuzulTicketCount" class="hidden">
                    <td>عدد التذاكر</td>
                    <td><input type="number" id="ticketCount" name="ticketCount" min="0"></td>
                </tr>
                <tr id="nuzulTicketTypes" class="hidden">
                    <td>أنواع التذاكر</td>
                    <td><input type="text" id="ticketTypes" name="ticketTypes"></td>
                </tr>
                <tr id="nuzulComplaints" class="hidden">
                    <td>الشكاوى</td>
                    <td><input type="number" id="nuzulComplaintsInput" name="nuzulComplaints" min="0"></td>
                </tr>
                <tr id="nuzulOpenTickets" class="hidden">
                    <td>التذاكر المفتوحة</td>
                    <td><input type="number" id="openTicket" name="openTicket" min="0"></td>
                </tr>
                <tr id="nuzulCleaning" class="hidden">
                    <td>التنظيف</td>
                    <td><input type="text" id="nuzulCleaning" name="nuzulCleaning"></td>
                </tr>
                <tr id="nuzulMaintenance" class="hidden">
                    <td>الصيانة</td>
                    <td><input type="text" id="nuzulMaintenance" name="nuzulMaintenance"></td>
                </tr>
                <tr id="nuzulPropertyOwners" class="hidden">
                    <td>أصحاب العقارات</td>
                    <td><input type="text" id="nuzulPropertyOwners" name="nuzulPropertyOwners"></td>
                </tr>
                <tr id="nuzulBooking" class="hidden">
                    <td>الحجز</td>
                    <td><input type="text" id="nuzulBooking" name="nuzulBooking"></td>
                </tr>
                <tr id="nuzulBookingClassification" class="hidden">
                    <td>تصنيف الحجز</td>
                    <td><input type="text" id="nuzulBookingClassification" name="nuzulBookingClassification"></td>
                </tr>
            </table>

            <div class="chart-container">
                <h3>أنواع التذاكر والشكاوى</h3>
                <canvas id="ticketsChart"></canvas>
            </div>

            <div class="chart-container">
                <h3>بيانات الطوابير</h3>
                <canvas id="queuesChart"></canvas>
            </div>

            <div class="chart-container hidden" id="bookingsChartContainer">
                <h3>بيانات الحجوزات</h3>
                <canvas id="bookingsChart"></canvas>
            </div>

            <div class="notes-section">
                <h3>الملاحظات</h3>
                <textarea id="notes" name="notes"></textarea>
            </div>

            <button type="submit">إرسال التقرير</button>
            <button type="button" id="exportPdf">تصدير إلى PDF</button>
            <button type="button" id="exportExcel">تصدير إلى Excel</button>
        </form>
    </div>
    <script>
        // Ensure libraries are loaded
        if (!window.jspdf || !window.Chart || !window.ExcelJS) {
            console.error('فشل تحميل إحدى المكتبات (jsPDF, Chart.js, ExcelJS).');
            alert('فشل تحميل المكتبات المطلوبة. يرجى التحقق من الاتصال بالإنترنت.');
        }

        const { jsPDF } = window.jspdf || {};
        const dateInput = document.getElementById('reportDate');
        const displayDate = document.getElementById('displayDate');
        const startDateInput = document.getElementById('startDate');
        const displayStartDate = document.getElementById('displayStartDate');
        const endDateInput = document.getElementById('endDate');
        const displayEndDate = document.getElementById('displayEndDate');
        const form = document.getElementById('reportForm');
        const projectSelect = document.getElementById('project');
        const queuesContainer = document.getElementById('queuesContainer');
        const queuesList = document.getElementById('queuesList');
        const reportTypeSelect = document.getElementById('reportType');
        const extraFields = document.getElementById('extraFields');
        const bookingsChartContainer = document.getElementById('bookingsChartContainer');
        const exportPdfButton = document.getElementById('exportPdf');
        const exportExcelButton = document.getElementById('exportExcel');
        const queuesChartCanvas = document.getElementById('queuesChart');
        const ticketsChartCanvas = document.getElementById('ticketsChart');
        const callsChartCanvas = document.getElementById('callsChart');
        const bookingsChartCanvas = document.getElementById('bookingsChart');

        let queuesChart, ticketsChart, callsChart, bookingsChart;

        const queues = {
            TMC: [
                'TMC Aramco AR',
                'TMC Aramco EN',
                'TMC Sahel AR',
                'TMC Sahel EN',
                'TMC Total AR',
                'TMC Total EN'
            ],
            NUZUL: [
                'نزل الحجوزات اليومية',
                'نزل محدود',
                'نزل محدود EN',
                'نزل الحجوزات الشهرية',
                'نزل علاقات الملاك',
                'نزل خدمات العملاء',
                'نزل صادر'
            ]
        };

        const chartColors = ['#e91e63', '#4caf50', '#9c27b0', '#ff9800', '#ff5722', '#8bc34a', '#f44336'];

        function formatDateWithDay(date) {
            const days = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
            const selectedDate = new Date(date);
            const dayName = days[selectedDate.getDay()];
            const day = String(selectedDate.getDate()).padStart(2, '0');
            const month = String(selectedDate.getMonth() + 1).padStart(2, '0');
            const year = selectedDate.getFullYear();
            return `${dayName} ${day}-${month}-${year}`;
        }

        function updateTicketFields() {
            const project = projectSelect.value;
            document.querySelectorAll('[id^="tmc"]').forEach(row => row.classList.add('hidden'));
            document.querySelectorAll('[id^="nuzul"]').forEach(row => row.classList.add('hidden'));
            if (project === 'TMC') {
                document.getElementById('tmcComplaints').classList.remove('hidden');
                document.getElementById('tmcContactRequests').classList.remove('hidden');
                document.getElementById('tmcExtraCleaning').classList.remove('hidden');
                document.getElementById('tmcExtraMaintenance').classList.remove('hidden');
                document.getElementById('tmcExtraPropertyOwners').classList.remove('hidden');
                document.getElementById('tmcExtraBooking').classList.remove('hidden');
                document.getElementById('tmcExtraBookingClassification').classList.remove('hidden');
                document.getElementById('tmcExtraMonthlyBooking').classList.remove('hidden');
                document.getElementById('tmcExtraDailyBooking').classList.remove('hidden');
                bookingsChartContainer.classList.add('hidden');
            } else if (project === 'NUZUL') {
                document.getElementById('nuzulTicketCount').classList.remove('hidden');
                document.getElementById('nuzulTicketTypes').classList.remove('hidden');
                document.getElementById('nuzulComplaints').classList.remove('hidden');
                document.getElementById('nuzulOpenTickets').classList.remove('hidden');
                document.getElementById('nuzulCleaning').classList.remove('hidden');
                document.getElementById('nuzulMaintenance').classList.remove('hidden');
                document.getElementById('nuzulPropertyOwners').classList.remove('hidden');
                document.getElementById('nuzulBooking').classList.remove('hidden');
                document.getElementById('nuzulBookingClassification').classList.remove('hidden');
                bookingsChartContainer.classList.remove('hidden');
            }
        }

        function updateCharts(reportData) {
            try {
                // Calls Chart
                const callData = [];
                const callLabels = ['المكالمات المستلمة', 'المكالمات المجابة', 'المكالمات غير المجابة', 'المكالمات المهجورة', 'المكالمات الصادرة'];
                const callFields = ['receivedCalls', 'answeredCalls', 'unansweredCalls', 'abandonedCalls', 'outgoingCalls'];
                callFields.forEach(field => {
                    callData.push(parseFloat(reportData[field]) || 0);
                });
                if (callsChart) callsChart.destroy();
                callsChart = new Chart(callsChartCanvas, {
                    type: 'doughnut',
                    data: {
                        labels: callLabels.filter((_, i) => callData[i] > 0),
                        datasets: [{
                            data: callData.filter(val => val > 0),
                            backgroundColor: chartColors.slice(0, callData.filter(val => val > 0).length),
                            borderWidth: 1
                        }]
                    },
                    options: {
                        responsive: true,
                        plugins: {
                            legend: { position: 'bottom' },
                            title: { display: true, text: 'توزيع إحصائيات المكالمات', font: { family: 'Amiri' } },
                            datalabels: {
                                color: '#fff',
                                formatter: (value, context) => {
                                    const total = context.dataset.data.reduce((sum, val) => sum + val, 0);
                                    return total ? `${((value / total) * 100).toFixed(1)}%` : '0%';
                                },
                                font: { weight: 'bold', size: 12, family: 'Amiri' }
                            }
                        },
                        cutout: '60%'
                    },
                    plugins: [ChartDataLabels]
                });

                // Queues Chart
                const queueData = [];
                const queueLabels = [];
                Object.keys(reportData).forEach(key => {
                    if (key.startsWith('queue_') && reportData[key]) {
                        queueLabels.push(key.replace('queue_', '').replace(/_/g, ' '));
                        queueData.push(parseFloat(reportData[key]) || 0);
                    }
                });
                if (queuesChart) queuesChart.destroy();
                queuesChart = new Chart(queuesChartCanvas, {
                    type: 'doughnut',
                    data: {
                        labels: queueLabels,
                        datasets: [{
                            data: queueData,
                            backgroundColor: chartColors.slice(0, queueData.length),
                            borderWidth: 1
                        }]
                    },
                    options: {
                        responsive: true,
                        plugins: {
                            legend: { position: 'bottom' },
                            title: { display: true, text: 'توزيع الطوابير', font: { family: 'Amiri' } },
                            datalabels: {
                                color: '#fff',
                                formatter: (value, context) => {
                                    const total = context.dataset.data.reduce((sum, val) => sum + val, 0);
                                    return total ? `${((value / total) * 100).toFixed(1)}%` : '0%';
                                },
                                font: { weight: 'bold', size: 12, family: 'Amiri' }
                            }
                        },
                        cutout: '60%'
                    },
                    plugins: [ChartDataLabels]
                });

                // Tickets Chart
                const ticketData = [];
                const ticketLabels = [];
                if (reportData.project === 'TMC') {
                    if (reportData.complaints) {
                        ticketLabels.push('الشكاوى');
                        ticketData.push(parseFloat(reportData.complaints) || 0);
                    }
                    if (reportData.contactRequestsInput) {
                        ticketLabels.push('طلبات التواصل');
                        ticketData.push(parseFloat(reportData.contactRequestsInput) || 0);
                    }
                } else if (reportData.project === 'NUZUL') {
                    if (reportData.ticketCount) {
                        ticketLabels.push('عدد التذاكر');
                        ticketData.push(parseFloat(reportData.ticketCount) || 0);
                    }
                    if (reportData.ticketTypes) {
                        reportData.ticketTypes.split(',').map(t => t.trim()).forEach(type => {
                            ticketLabels.push(`نوع التذكرة: ${type}`);
                            ticketData.push(1);
                        });
                    }
                    if (reportData.nuzulComplaints) {
                        ticketLabels.push('الشكاوى');
                        ticketData.push(parseFloat(reportData.nuzulComplaints) || 0);
                    }
                    if (reportData.openTicket) {
                        ticketLabels.push('التذاكر المفتوحة');
                        ticketData.push(parseFloat(reportData.openTicket) || 0);
                    }
                    if (reportData.nuzulCleaning) {
                        ticketLabels.push('التنظيف');
                        ticketData.push(parseFloat(reportData.nuzulCleaning) || 1);
                    }
                    if (reportData.nuzulMaintenance) {
                        ticketLabels.push('الصيانة');
                        ticketData.push(parseFloat(reportData.nuzulMaintenance) || 1);
                    }
                    if (reportData.nuzulPropertyOwners) {
                        ticketLabels.push('أصحاب العقارات');
                        ticketData.push(parseFloat(reportData.nuzulPropertyOwners) || 1);
                    }
                    if (reportData.nuzulBooking) {
                        ticketLabels.push('الحجز');
                        ticketData.push(parseFloat(reportData.nuzulBooking) || 1);
                    }
                    if (reportData.nuzulBookingClassification) {
                        ticketLabels.push('تصنيف الحجز');
                        ticketData.push(parseFloat(reportData.nuzulBookingClassification) || 1);
                    }
                }
                if (ticketsChart) ticketsChart.destroy();
                ticketsChart = new Chart(ticketsChartCanvas, {
                    type: 'doughnut',
                    data: {
                        labels: ticketLabels,
                        datasets: [{
                            data: ticketData,
                            backgroundColor: chartColors.slice(0, ticketData.length),
                            borderWidth: 1
                        }]
                    },
                    options: {
                        responsive: true,
                        plugins: {
                            legend: { position: 'bottom' },
                            title: { display: true, text: 'توزيع أنواع التذاكر والشكاوى', font: { family: 'Amiri' } },
                            datalabels: {
                                color: '#fff',
                                formatter: (value, context) => {
                                    const total = context.dataset.data.reduce((sum, val) => sum + val, 0);
                                    return total ? `${((value / total) * 100).toFixed(1)}%` : '0%';
                                },
                                font: { weight: 'bold', size: 12, family: 'Amiri' }
                            }
                        },
                        cutout: '60%'
                    },
                    plugins: [ChartDataLabels]
                });

                // Bookings Chart (NUZUL only)
                if (reportData.project === 'NUZUL') {
                    const bookingData = [];
                    const bookingLabels = [];
                    if (reportData.dailyBooking) {
                        bookingLabels.push('الحجوزات اليومية');
                        bookingData.push(parseFloat(reportData.dailyBooking) || 0);
                    }
                    if (reportData.monthlyBooking) {
                        bookingLabels.push('الحجوزات الشهرية');
                        bookingData.push(parseFloat(reportData.monthlyBooking) || 0);
                    }
                    if (bookingsChart) bookingsChart.destroy();
                    bookingsChart = new Chart(bookingsChartCanvas, {
                        type: 'doughnut',
                        data: {
                            labels: bookingLabels,
                            datasets: [{
                                data: bookingData,
                                backgroundColor: chartColors.slice(0, bookingData.length),
                                borderWidth: 1
                            }]
                        },
                        options: {
                            responsive: true,
                            plugins: {
                                legend: { position: 'bottom' },
                                title: { display: true, text: 'توزيع الحجوزات', font: { family: 'Amiri' } },
                                datalabels: {
                                    color: '#fff',
                                    formatter: (value, context) => {
                                        const total = context.dataset.data.reduce((sum, val) => sum + val, 0);
                                        return total ? `${((value / total) * 100).toFixed(1)}%` : '0%';
                                    },
                                    font: { weight: 'bold', size: 12, family: 'Amiri' }
                                }
                            },
                            cutout: '60%'
                        },
                        plugins: [ChartDataLabels]
                    });
                } else {
                    if (bookingsChart) bookingsChart.destroy();
                    bookingsChart = null;
                }
            } catch (e) {
                console.error('خطأ في تحديث الرسوم البيانية:', e);
                alert('فشل في تحديث الرسوم البيانية. تحقق من وحدة التحكم.');
            }
        }

        function collectReportData() {
            try {
                const reportData = {
                    reportType: document.getElementById('reportType').value,
                    project: document.getElementById('project').value,
                    department: document.getElementById('department').value,
                    receivedCalls: document.getElementById('receivedCalls').value,
                    answeredCalls: document.getElementById('answeredCalls').value,
                    unansweredCalls: document.getElementById('unansweredCalls').value,
                    abandonedCalls: document.getElementById('abandonedCalls').value,
                    outgoingCalls: document.getElementById('outgoingCalls').value,
                    notes: document.getElementById('notes').value
                };

                if (reportData.project === 'TMC') {
                    reportData.complaints = document.getElementById('complaints').value;
                    reportData.contactRequestsInput = document.getElementById('contactRequestsInput').value;
                    reportData.cleaning = document.getElementById('cleaning').value;
                    reportData.maintenance = document.getElementById('maintenance').value;
                    reportData.propertyOwners = document.getElementById('propertyOwners').value;
                    reportData.booking = document.getElementById('booking').value;
                    reportData.bookingClassification = document.getElementById('bookingClassification').value;
                    reportData.monthlyBooking = document.getElementById('monthlyBooking').value;
                    reportData.dailyBooking = document.getElementById('dailyBooking').value;
                } else if (reportData.project === 'NUZUL') {
                    reportData.ticketCount = document.getElementById('ticketCount').value;
                    reportData.ticketTypes = document.getElementById('ticketTypes').value;
                    reportData.nuzulComplaints = document.getElementById('nuzulComplaintsInput').value;
                    reportData.openTicket = document.getElementById('openTicket').value;
                    reportData.nuzulCleaning = document.getElementById('nuzulCleaning').value;
                    reportData.nuzulMaintenance = document.getElementById('nuzulMaintenance').value;
                    reportData.nuzulPropertyOwners = document.getElementById('nuzulPropertyOwners').value;
                    reportData.nuzulBooking = document.getElementById('nuzulBooking').value;
                    reportData.nuzulBookingClassification = document.getElementById('nuzulBookingClassification').value;
                    reportData.dailyBooking = document.getElementById('dailyBooking').value;
                    reportData.monthlyBooking = document.getElementById('monthlyBooking').value;
                }

                if (reportData.reportType === 'Daily') {
                    reportData.date = displayDate.textContent;
                } else {
                    reportData.startDate = displayStartDate.textContent;
                    reportData.endDate = displayEndDate.textContent;
                }

                const queueInputs = queuesList.querySelectorAll('input');
                queueInputs.forEach(input => {
                    if (input.value) {
                        reportData[input.name] = input.value;
                    }
                });

                if (!extraFields.classList.contains('hidden')) {
                    reportData.avgCallDuration = document.getElementById('avgCallDuration').value;
                    reportData.avgWaitTime = document.getElementById('avgWaitTime').value;
                    const peakStart = document.getElementById('peakStart').value;
                    const peakEnd = document.getElementById('peakEnd').value;
                    if (peakStart && peakEnd) {
                        reportData.peakTimes = `من ${peakStart} إلى ${peakEnd}`;
                    }
                }

                return reportData;
            } catch (e) {
                console.error('خطأ في جمع بيانات التقرير:', e);
                alert('فشل في جمع بيانات التقرير. تحقق من وحدة التحكم.');
                return null;
            }
        }

        dateInput.addEventListener('change', function() {
            displayDate.textContent = dateInput.value ? formatDateWithDay(dateInput.value) : '';
        });

        startDateInput.addEventListener('change', function() {
            displayStartDate.textContent = startDateInput.value ? formatDateWithDay(startDateInput.value) : '';
        });

        endDateInput.addEventListener('change', function() {
            displayEndDate.textContent = endDateInput.value ? formatDateWithDay(endDateInput.value) : '';
        });

        // Set default date
        const today = new Date();
        dateInput.value = today.toISOString().split('T')[0];
        displayDate.textContent = formatDateWithDay(today);

        projectSelect.addEventListener('change', function() {
            const selectedProject = projectSelect.value;
            queuesList.innerHTML = '';
            if (selectedProject && queues[selectedProject]) {
                queuesContainer.classList.remove('hidden');
                let tableContent = '<table><tr><th>الطابور</th><th>البيانات</th></tr>';
                queues[selectedProject].forEach(queue => {
                    tableContent += `<tr><td>${queue}</td><td><input type="number" name="queue_${queue.replace(/\s/g, '_')}" placeholder="بيانات ${queue}" min="0"></td></tr>`;
                });
                tableContent += '</table>';
                queuesList.innerHTML = tableContent;
            } else {
                queuesContainer.classList.add('hidden');
            }
            updateTicketFields();
            const reportData = collectReportData();
            if (reportData) updateCharts(reportData);
        });

        reportTypeSelect.addEventListener('change', function() {
            const reportType = reportTypeSelect.value;
            const project = projectSelect.value;
            if (reportType === 'Daily') {
                document.getElementById('singleDateGroup').classList.remove('hidden');
                document.getElementById('dateRangeGroup').classList.add('hidden');
                dateInput.required = true;
                startDateInput.required = false;
                endDateInput.required = false;
                extraFields.classList.add('hidden');
            } else {
                document.getElementById('singleDateGroup').classList.add('hidden');
                document.getElementById('dateRangeGroup').classList.remove('hidden');
                dateInput.required = false;
                startDateInput.required = true;
                endDateInput.required = true;
                extraFields.classList.remove('hidden');
                document.querySelectorAll('[id^="tmcExtra"]').forEach(row => {
                    row.classList.toggle('hidden', project !== 'TMC');
                });
            }
        });

        async function captureChart(canvas) {
            return new Promise(resolve => {
                setTimeout(() => {
                    try {
                        resolve(canvas.toDataURL('image/png'));
                    } catch (e) {
                        console.error('خطأ في التقاط الرسم البياني:', e);
                        resolve(null);
                    }
                }, 4000);
            });
        }

        async function generatePDF(reportData) {
            if (!jsPDF) {
                alert('فشل تحميل مكتبة jsPDF. تحقق من الاتصال بالإنترنت.');
                return;
            }
            try {
                const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
                try {
                    doc.addFont('https://cdn.jsdelivr.net/npm/amiri-font@0.117.0/Amiri-Regular.ttf', 'Amiri', 'normal');
                    doc.setFont('Amiri');
                } catch (e) {
                    console.warn('فشل تحميل خط Amiri، يتم استخدام helvetica:', e);
                    doc.setFont('helvetica');
                }

                doc.setFontSize(16);
                doc.setTextColor(0, 188, 212);
                doc.text(`${reportData.project || 'مركز الاتصال'} تقرير - ${reportData.reportType}`, 190, 20, { align: 'right' });
                doc.setFontSize(12);

                let y = 30;
                if (reportData.date) {
                    doc.text(`التاريخ: ${reportData.date}`, 190, y, { align: 'right' });
                    y += 10;
                } else if (reportData.startDate && reportData.endDate) {
                    doc.text(`نطاق التاريخ: ${reportData.startDate} إلى ${reportData.endDate}`, 190, y, { align: 'right' });
                    y += 10;
                }

                const [callChartImg, ticketChartImg, queueChartImg, bookingsChartImg] = await Promise.all([
                    captureChart(callsChartCanvas),
                    captureChart(ticketsChartCanvas),
                    captureChart(queuesChartCanvas),
                    reportData.project === 'NUZUL' ? captureChart(bookingsChartCanvas) : Promise.resolve(null)
                ]);

                // Call Statistics
                const callStats = [];
                if (reportData.receivedCalls) callStats.push(['المكالمات المستلمة', reportData.receivedCalls]);
                if (reportData.answeredCalls) callStats.push(['المكالمات المجابة', reportData.answeredCalls]);
                if (reportData.unansweredCalls) callStats.push(['المكالمات غير المجابة', reportData.unansweredCalls]);
                if (reportData.abandonedCalls) callStats.push(['المكالمات المهجورة', reportData.abandonedCalls]);
                if (reportData.outgoingCalls) callStats.push(['المكالمات الصادرة', reportData.outgoingCalls]);
                if (callStats.length > 0) {
                    doc.text('إحصائيات المكالمات', 190, y, { align: 'right' });
                    y += 5;
                    doc.autoTable({
                        startY: y,
                        head: [['الحقل', 'القيمة']],
                        body: callStats,
                        theme: 'grid',
                        styles: { font: 'Amiri', halign: 'right', cellPadding: 2 },
                        headStyles: { fillColor: [224, 247, 250], textColor: [0, 121, 107], font: 'Amiri', halign: 'right' },
                        margin: { right: 105 }
                    });
                    const callTableHeight = doc.lastAutoTable.finalY - y;
                    if (callChartImg) doc.addImage(callChartImg, 'PNG', 20, y + (callTableHeight - 50) / 2, 50, 50);
                    y = doc.lastAutoTable.finalY + 10;
                }

                // Performance Metrics
                const performanceMetrics = [];
                if (reportData.avgCallDuration) performanceMetrics.push(['متوسط مدة المكالمة', reportData.avgCallDuration]);
                if (reportData.avgWaitTime) performanceMetrics.push(['متوسط وقت الانتظار', reportData.avgWaitTime]);
                if (reportData.peakTimes) performanceMetrics.push(['أوقات الذروة', reportData.peakTimes]);
                if (reportData.project === 'TMC') {
                    if (reportData.cleaning) performanceMetrics.push(['التنظيف', reportData.cleaning]);
                    if (reportData.maintenance) performanceMetrics.push(['الصيانة', reportData.maintenance]);
                    if (reportData.propertyOwners) performanceMetrics.push(['أصحاب العقارات', reportData.propertyOwners]);
                    if (reportData.booking) performanceMetrics.push(['الحجز', reportData.booking]);
                    if (reportData.bookingClassification) performanceMetrics.push(['تصنيف الحجز', reportData.bookingClassification]);
                    if (reportData.monthlyBooking) performanceMetrics.push(['الحجوزات الشهرية', reportData.monthlyBooking]);
                    if (reportData.dailyBooking) performanceMetrics.push(['الحجوزات اليومية', reportData.dailyBooking]);
                }
                if (performanceMetrics.length > 0) {
                    doc.text('مؤشرات الأداء', 190, y, { align: 'right' });
                    y += 5;
                    doc.autoTable({
                        startY: y,
                        head: [['الحقل', 'القيمة']],
                        body: performanceMetrics,
                        theme: 'grid',
                        styles: { font: 'Amiri', halign: 'right', cellPadding: 2 },
                        headStyles: { fillColor: [224, 247, 250], textColor: [0, 121, 107], font: 'Amiri', halign: 'right' },
                        margin: { right: 105 }
                    });
                    y = doc.lastAutoTable.finalY + 10;
                }

                // Tickets
                const ticketData = [];
                if (reportData.project === 'TMC') {
                    if (reportData.complaints) ticketData.push(['الشكاوى', reportData.complaints]);
                    if (reportData.contactRequestsInput) ticketData.push(['طلبات التواصل', reportData.contactRequestsInput]);
                } else if (reportData.project === 'NUZUL') {
                    if (reportData.ticketCount) ticketData.push(['عدد التذاكر', reportData.ticketCount]);
                    if (reportData.openTicket) ticketData.push(['التذاكر المفتوحة', reportData.openTicket]);
                    if (reportData.ticketTypes) {
                        reportData.ticketTypes.split(',').map(t => t.trim()).forEach(type => {
                            ticketData.push(['نوع التذكرة', type]);
                        });
                    }
                    if (reportData.nuzulComplaints) ticketData.push(['الشكاوى', reportData.nuzulComplaints]);
                    if (reportData.nuzulCleaning) ticketData.push(['التنظيف', reportData.nuzulCleaning]);
                    if (reportData.nuzulMaintenance) ticketData.push(['الصيانة', reportData.nuzulMaintenance]);
                    if (reportData.nuzulPropertyOwners) ticketData.push(['أصحاب العقارات', reportData.nuzulPropertyOwners]);
                    if (reportData.nuzulBooking) ticketData.push(['الحجز', reportData.nuzulBooking]);
                    if (reportData.nuzulBookingClassification) ticketData.push(['تصنيف الحجز', reportData.nuzulBookingClassification]);
                }
                if (ticketData.length > 0) {
                    doc.text('التذاكر', 190, y, { align: 'right' });
                    y += 5;
                    doc.autoTable({
                        startY: y,
                        head: [['الحقل', 'القيمة']],
                        body: ticketData,
                        theme: 'grid',
                        styles: { font: 'Amiri', halign: 'right', cellPadding: 2 },
                        headStyles: { fillColor: [224, 247, 250], textColor: [0, 121, 107], font: 'Amiri', halign: 'right' },
                        margin: { right: 105 }
                    });
                    const ticketTableHeight = doc.lastAutoTable.finalY - y;
                    if (ticketChartImg) doc.addImage(ticketChartImg, 'PNG', 20, y + (ticketTableHeight - 50) / 2, 50, 50);
                    y = doc.lastAutoTable.finalY + 10;
                }

                // Queues
                const queueData = [];
                Object.keys(reportData).forEach(key => {
                    if (key.startsWith('queue_') && reportData[key]) {
                        queueData.push([key.replace('queue_', '').replace(/_/g, ' '), reportData[key]]);
                    }
                });
                if (queueData.length > 0) {
                    doc.text('الطوابير', 190, y, { align: 'right' });
                    y += 5;
                    doc.autoTable({
                        startY: y,
                        head: [['الطابور', 'البيانات']],
                        body: queueData,
                        theme: 'grid',
                        styles: { font: 'Amiri', halign: 'right', cellPadding: 2 },
                        headStyles: { fillColor: [224, 247, 250], textColor: [0, 121, 107], font: 'Amiri', halign: 'right' },
                        margin: { right: 105 }
                    });
                    const queueTableHeight = doc.lastAutoTable.finalY - y;
                    if (queueChartImg) doc.addImage(queueChartImg, 'PNG', 20, y + (queueTableHeight - 50) / 2, 50, 50);
                    y = doc.lastAutoTable.finalY + 10;
                }

                // Bookings (NUZUL only)
                if (reportData.project === 'NUZUL') {
                    const bookingData = [];
                    if (reportData.dailyBooking) bookingData.push(['الحجوزات اليومية', reportData.dailyBooking]);
                    if (reportData.monthlyBooking) bookingData.push(['الحجوزات الشهرية', reportData.monthlyBooking]);
                    if (bookingData.length > 0) {
                        doc.text('الحجوزات', 190, y, { align: 'right' });
                        y += 5;
                        doc.autoTable({
                            startY: y,
                            head: [['الحقل', 'القيمة']],
                            body: bookingData,
                            theme: 'grid',
                            styles: { font: 'Amiri', halign: 'right', cellPadding: 2 },
                            headStyles: { fillColor: [224, 247, 250], textColor: [0, 121, 107], font: 'Amiri', halign: 'right' },
                            margin: { right: 105 }
                        });
                        const bookingTableHeight = doc.lastAutoTable.finalY - y;
                        if (bookingsChartImg) doc.addImage(bookingsChartImg, 'PNG', 20, y + (bookingTableHeight - 50) / 2, 50, 50);
                        y = doc.lastAutoTable.finalY + 10;
                    }
                }

                // Notes
                if (reportData.notes && reportData.notes.trim()) {
                    doc.text('الملاحظات', 190, y, { align: 'right' });
                    y += 7;
                    const lines = doc.splitTextToSize(reportData.notes, 170);
                    lines.forEach(line => {
                        doc.text(line, 190, y, { align: 'right' });
                        y += 7;
                    });
                }

                doc.save(`تقرير_مركز_الاتصال_${new Date().toISOString().split('T')[0]}.pdf`);
            } catch (e) {
                console.error('خطأ في إنشاء PDF:', e);
                alert('فشل في تصدير PDF. تحقق من وحدة التحكم.');
            }
        }

        async function generateExcel(reportData) {
            if (!window.ExcelJS) {
                alert('فشل تحميل مكتبة ExcelJS. تحقق من الاتصال بالإنترنت.');
                return;
            }
            try {
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet('تقرير مركز الاتصال');
                worksheet.columns = [{ width: 30 }, { width: 20 }];

                let rowIndex = 1;
                worksheet.addRow([`${reportData.project || 'مركز الاتصال'} تقرير - ${reportData.reportType}`]);
                worksheet.getRow(rowIndex).font = { size: 16, bold: true, color: { argb: '00BCD4' } };
                worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                rowIndex += 2;

                // Basic Information
                const basicInfo = [
                    ['نوع التقرير', reportData.reportType],
                    ['المشروع', reportData.project],
                    ['القسم', reportData.department],
                    ...(reportData.date ? [['التاريخ', reportData.date]] : []),
                    ...(reportData.startDate ? [['تاريخ البداية', reportData.startDate]] : []),
                    ...(reportData.endDate ? [['تاريخ النهاية', reportData.endDate]] : [])
                ];
                if (basicInfo.length > 0) {
                    worksheet.addRow(['المعلومات الأساسية']);
                    worksheet.getRow(rowIndex).font = { size: 12, bold: true };
                    worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                    rowIndex++;
                    worksheet.addTable({
                        name: 'BasicInformation',
                        ref: `A${rowIndex}`,
                        headerRow: true,
                        columns: [{ name: 'الحقل' }, { name: 'القيمة' }],
                        rows: basicInfo
                    });
                    const basicTable = worksheet.getTable('BasicInformation');
                    basicTable.tableRef.split(':').forEach(cell => {
                        worksheet.getCell(cell).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                        worksheet.getCell(cell).alignment = { horizontal: 'right' };
                    });
                    rowIndex = parseInt(basicTable.tableRef.split(':')[1].match(/\d+/)[0]) + 2;
                }

                // Call Statistics
                const callStats = [];
                if (reportData.receivedCalls) callStats.push(['المكالمات المستلمة', reportData.receivedCalls]);
                if (reportData.answeredCalls) callStats.push(['المكالمات المجابة', reportData.answeredCalls]);
                if (reportData.unansweredCalls) callStats.push(['المكالمات غير المجابة', reportData.unansweredCalls]);
                if (reportData.abandonedCalls) callStats.push(['المكالمات المهجورة', reportData.abandonedCalls]);
                if (reportData.outgoingCalls) callStats.push(['المكالمات الصادرة', reportData.outgoingCalls]);
                if (callStats.length > 0) {
                    worksheet.addRow(['إحصائيات المكالمات']);
                    worksheet.getRow(rowIndex).font = { size: 12, bold: true };
                    worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                    rowIndex++;
                    worksheet.addTable({
                        name: 'CallStatistics',
                        ref: `A${rowIndex}`,
                        headerRow: true,
                        columns: [{ name: 'الحقل' }, { name: 'القيمة' }],
                        rows: callStats
                    });
                    const callTable = worksheet.getTable('CallStatistics');
                    callTable.tableRef.split(':').forEach(cell => {
                        worksheet.getCell(cell).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                        worksheet.getCell(cell).alignment = { horizontal: 'right' };
                    });
                    const callTableEndRow = parseInt(callTable.tableRef.split(':')[1].match(/\d+/)[0]);
                    const callChartImg = await captureChart(callsChartCanvas);
                    if (callChartImg) {
                        const imageId = workbook.addImage({ base64: callChartImg, extension: 'png' });
                        worksheet.addImage(imageId, { tl: { col: 2.5, row: rowIndex - 1 }, ext: { width: 200, height: 200 } });
                    }
                    rowIndex = callTableEndRow + 2;
                }

                // Performance Metrics
                const performanceMetrics = [];
                if (reportData.avgCallDuration) performanceMetrics.push(['متوسط مدة المكالمة', reportData.avgCallDuration]);
                if (reportData.avgWaitTime) performanceMetrics.push(['متوسط وقت الانتظار', reportData.avgWaitTime]);
                if (reportData.peakTimes) performanceMetrics.push(['أوقات الذروة', reportData.peakTimes]);
                if (reportData.project === 'TMC') {
                    if (reportData.cleaning) performanceMetrics.push(['التنظيف', reportData.cleaning]);
                    if (reportData.maintenance) performanceMetrics.push(['الصيانة', reportData.maintenance]);
                    if (reportData.propertyOwners) performanceMetrics.push(['أصحاب العقارات', reportData.propertyOwners]);
                    if (reportData.booking) performanceMetrics.push(['الحجز', reportData.booking]);
                    if (reportData.bookingClassification) performanceMetrics.push(['تصنيف الحجز', reportData.bookingClassification]);
                    if (reportData.monthlyBooking) performanceMetrics.push(['الحجوزات الشهرية', reportData.monthlyBooking]);
                    if (reportData.dailyBooking) performanceMetrics.push(['الحجوزات اليومية', reportData.dailyBooking]);
                }
                if (performanceMetrics.length > 0) {
                    worksheet.addRow(['مؤشرات الأداء']);
                    worksheet.getRow(rowIndex).font = { size: 12, bold: true };
                    worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                    rowIndex++;
                    worksheet.addTable({
                        name: 'PerformanceMetrics',
                        ref: `A${rowIndex}`,
                        headerRow: true,
                        columns: [{ name: 'الحقل' }, { name: 'القيمة' }],
                        rows: performanceMetrics
                    });
                    const perfTable = worksheet.getTable('PerformanceMetrics');
                    perfTable.tableRef.split(':').forEach(cell => {
                        worksheet.getCell(cell).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                        worksheet.getCell(cell).alignment = { horizontal: 'right' };
                    });
                    rowIndex = parseInt(perfTable.tableRef.split(':')[1].match(/\d+/)[0]) + 2;
                }

                // Tickets
                const ticketData = [];
                if (reportData.project === 'TMC') {
                    if (reportData.complaints) ticketData.push(['الشكاوى', reportData.complaints]);
                    if (reportData.contactRequestsInput) ticketData.push(['طلبات التواصل', reportData.contactRequestsInput]);
                } else if (reportData.project === 'NUZUL') {
                    if (reportData.ticketCount) ticketData.push(['عدد التذاكر', reportData.ticketCount]);
                    if (reportData.openTicket) ticketData.push(['التذاكر المفتوحة', reportData.openTicket]);
                    if (reportData.ticketTypes) {
                        reportData.ticketTypes.split(',').map(t => t.trim()).forEach(type => {
                            ticketData.push(['نوع التذكرة', type]);
                        });
                    }
                    if (reportData.nuzulComplaints) ticketData.push(['الشكاوى', reportData.nuzulComplaints]);
                    if (reportData.nuzulCleaning) ticketData.push(['التنظيف', reportData.nuzulCleaning]);
                    if (reportData.nuzulMaintenance) ticketData.push(['الصيانة', reportData.nuzulMaintenance]);
                    if (reportData.nuzulPropertyOwners) ticketData.push(['أصحاب العقارات', reportData.nuzulPropertyOwners]);
                    if (reportData.nuzulBooking) ticketData.push(['الحجز', reportData.nuzulBooking]);
                    if (reportData.nuzulBookingClassification) ticketData.push(['تصنيف الحجز', reportData.nuzulBookingClassification]);
                }
                if (ticketData.length > 0) {
                    worksheet.addRow(['التذاكر']);
                    worksheet.getRow(rowIndex).font = { size: 12, bold: true };
                    worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                    rowIndex++;
                    worksheet.addTable({
                        name: 'Tickets',
                        ref: `A${rowIndex}`,
                        headerRow: true,
                        columns: [{ name: 'الحقل' }, { name: 'القيمة' }],
                        rows: ticketData
                    });
                    const ticketTable = worksheet.getTable('Tickets');
                    ticketTable.tableRef.split(':').forEach(cell => {
                        worksheet.getCell(cell).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                        worksheet.getCell(cell).alignment = { horizontal: 'right' };
                    });
                    const ticketTableEndRow = parseInt(ticketTable.tableRef.split(':')[1].match(/\d+/)[0]);
                    const ticketChartImg = await captureChart(ticketsChartCanvas);
                    if (ticketChartImg) {
                        const imageId = workbook.addImage({ base64: ticketChartImg, extension: 'png' });
                        worksheet.addImage(imageId, { tl: { col: 2.5, row: rowIndex - 1 }, ext: { width: 200, height: 200 } });
                    }
                    rowIndex = ticketTableEndRow + 2;
                }

                // Queues
                const queueData = [];
                Object.keys(reportData).forEach(key => {
                    if (key.startsWith('queue_') && reportData[key]) {
                        queueData.push([key.replace('queue_', '').replace(/_/g, ' '), reportData[key]]);
                    }
                });
                if (queueData.length > 0) {
                    worksheet.addRow(['الطوابير']);
                    worksheet.getRow(rowIndex).font = { size: 12, bold: true };
                    worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                    rowIndex++;
                    worksheet.addTable({
                        name: 'Queues',
                        ref: `A${rowIndex}`,
                        headerRow: true,
                        columns: [{ name: 'الطابور' }, { name: 'البيانات' }],
                        rows: queueData
                    });
                    const queueTable = worksheet.getTable('Queues');
                    queueTable.tableRef.split(':').forEach(cell => {
                        worksheet.getCell(cell).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                        worksheet.getCell(cell).alignment = { horizontal: 'right' };
                    });
                    const queueTableEndRow = parseInt(queueTable.tableRef.split(':')[1].match(/\d+/)[0]);
                    const queueChartImg = await captureChart(queuesChartCanvas);
                    if (queueChartImg) {
                        const imageId = workbook.addImage({ base64: queueChartImg, extension: 'png' });
                        worksheet.addImage(imageId, { tl: { col: 2.5, row: rowIndex - 1 }, ext: { width: 200, height: 200 } });
                    }
                    rowIndex = queueTableEndRow + 2;
                }

                // Bookings (NUZUL only)
                if (reportData.project === 'NUZUL') {
                    const bookingData = [];
                    if (reportData.dailyBooking) bookingData.push(['الحجوزات اليومية', reportData.dailyBooking]);
                    if (reportData.monthlyBooking) bookingData.push(['الحجوزات الشهرية', reportData.monthlyBooking]);
                    if (bookingData.length > 0) {
                        worksheet.addRow(['الحجوزات']);
                        worksheet.getRow(rowIndex).font = { size: 12, bold: true };
                        worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                        rowIndex++;
                        worksheet.addTable({
                            name: 'Bookings',
                            ref: `A${rowIndex}`,
                            headerRow: true,
                            columns: [{ name: 'الحقل' }, { name: 'القيمة' }],
                            rows: bookingData
                        });
                        const bookingTable = worksheet.getTable('Bookings');
                        bookingTable.tableRef.split(':').forEach(cell => {
                            worksheet.getCell(cell).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                            worksheet.getCell(cell).alignment = { horizontal: 'right' };
                        });
                        const bookingTableEndRow = parseInt(bookingTable.tableRef.split(':')[1].match(/\d+/)[0]);
                        const bookingChartImg = await captureChart(bookingsChartCanvas);
                        if (bookingChartImg) {
                            const imageId = workbook.addImage({ base64: bookingChartImg, extension: 'png' });
                            worksheet.addImage(imageId, { tl: { col: 2.5, row: rowIndex - 1 }, ext: { width: 200, height: 200 } });
                        }
                        rowIndex = bookingTableEndRow + 2;
                    }
                }

                // Notes
                if (reportData.notes && reportData.notes.trim()) {
                    worksheet.addRow(['الملاحظات']);
                    worksheet.getRow(rowIndex).font = { size: 12, bold: true };
                    worksheet.getRow(rowIndex).alignment = { horizontal: 'right' };
                    rowIndex++;
                    worksheet.addRow([reportData.notes]);
                    worksheet.getRow(rowIndex).alignment = { horizontal: 'right', wrapText: true };
                    worksheet.getRow(rowIndex).height = 100;
                }

                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `تقرير_مركز_الاتصال_${new Date().toISOString().split('T')[0]}.xlsx`;
                a.click();
                URL.revokeObjectURL(url);
            } catch (e) {
                console.error('خطأ في إنشاء Excel:', e);
                alert('فشل في تصدير Excel. تحقق من وحدة التحكم.');
            }
        }

        form.addEventListener('submit', function(e) {
            e.preventDefault();
            const reportData = collectReportData();
            if (reportData) {
                updateCharts(reportData);
                alert('تم إرسال التقرير بنجاح!');
            }
        });

        exportPdfButton.addEventListener('click', function() {
            const reportData = collectReportData();
            if (reportData) generatePDF(reportData);
        });

        exportExcelButton.addEventListener('click', function() {
            const reportData = collectReportData();
            if (reportData) generateExcel(reportData);
        });
    </script>
</body>
</html>
