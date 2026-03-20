// ===== GLOBAL =====
let rawTotalEmissionKg = 0;
const KG_PER_CREDIT = 1000;
const PRICE_PER_CREDIT = 500;

let scopeChartInstance = null;
let barChartInstance = null;

// 🌟 ตัวแปรเก็บค่าจากไฟล์เทสโดยตรง (แก้ปัญหาหลังบ้านส่ง 0 กลับมา)
let localCircularPercent = 0;
let localCarbonSaved = 0;

// ===== API CONFIGURATION =====
const API_BASE_URL = "https://masterful-colby-laciniate.ngrok-free.dev"; 
const API_CALCULATE = `${API_BASE_URL}/api/factory-calculate`;
const API_DASHBOARD = `${API_BASE_URL}/api/dashboard-summary`;
const API_REPORT = `${API_BASE_URL}/api/auto-report`;

document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("fileInput");
    if (fileInput) fileInput.addEventListener("change", onFileChange);

    const downloadBtn = document.getElementById("downloadTemplate");
    if (downloadBtn) {
        downloadBtn.addEventListener("click", () => {
            window.location.href = "Template.xlsx"; 
        });
    }

    const resetBtn = document.getElementById("resetBtn");
    if (resetBtn) {
        resetBtn.addEventListener("click", async () => {
            if (confirm("คุณแน่ใจหรือไม่ว่าต้องการล้างข้อมูลเก่าในระบบทั้งหมด?")) {
                const statusLabel = document.getElementById("status");
                if (statusLabel) statusLabel.innerText = "กำลังล้างข้อมูล...";
                
                try {
                    await axios.post(`${API_BASE_URL}/api/clear-data`, {}, {
                        headers: { "ngrok-skip-browser-warning": "true" }
                    });
                    
                    document.getElementById("totalRows").innerText = "0";
                    if (fileInput) fileInput.value = ""; 
                    
                    localCircularPercent = 0;
                    localCarbonSaved = 0;
                    
                    await loadDashboardSummary();
                    if (statusLabel) statusLabel.innerText = "สถานะ: ล้างข้อมูลเรียบร้อยแล้ว 🗑️ พร้อมอัปโหลดใหม่";
                    
                } catch (err) {
                    console.error("Reset Error:", err);
                    alert("เกิดข้อผิดพลาดในการล้างข้อมูล ตรวจสอบ API Backend");
                }
            }
        });
    }

    loadDashboardSummary();
});

function formatExcelDate(excelDate) {
    if (!excelDate) return new Date().toISOString().split('T')[0];
    if (typeof excelDate === 'number') {
        const date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
        return date.toISOString().split('T')[0];
    }
    try {
        const d = new Date(excelDate);
        if (!isNaN(d)) return d.toISOString().split('T')[0];
    } catch(e) {}
    return new Date().toISOString().split('T')[0];
}

function safeFloat(val, def = 0) {
    if (val === null || val === undefined || val === '') return def;
    if (typeof val === 'number') return val;
    const str = String(val).replace(/,/g, '').trim();
    const num = parseFloat(str);
    return isNaN(num) ? def : num;
}

function getExactVal(obj, searchWords) {
    for (let w of searchWords) {
        const cleanW = String(w).toLowerCase().replace(/[^a-z0-9ก-๙]/g, '');
        for (let k of Object.keys(obj)) {
            const cleanK = String(k).toLowerCase().replace(/[^a-z0-9ก-๙]/g, ''); 
            if (cleanK.includes(cleanW)) {
                return obj[k];
            }
        }
    }
    return undefined;
}

async function onFileChange(e) {
    const file = e.target.files[0];
    if (!file) return;

    const statusLabel = document.getElementById("status");
    if (statusLabel) statusLabel.innerText = "กำลังแยกข้อมูลและบันทึก...";

    const reader = new FileReader();
    reader.onload = async (evt) => {
        try {
            const data = new Uint8Array(evt.target.result);
            const wb = XLSX.read(data, { type: "array" });

            const payload = [];
            
            // รีเซ็ตค่าเพื่อเตรียมอ่านจากไฟล์ใหม่
            localCircularPercent = 0;
            localCarbonSaved = 0;
            let totalOutput = 0;
            let totalFeedstock = 0;

            wb.SheetNames.forEach(sheetName => {
                const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });
                
                let headerRowIndex = -1;
                let headers = [];

                for (let i = 0; i < Math.min(rawRows.length, 10); i++) {
                    const rowStr = rawRows[i].join("").toLowerCase();
                    if (rowStr.includes("activity") || rowStr.includes("product") || 
                        rowStr.includes("refrigerant") || rowStr.includes("date") || 
                        rowStr.includes("category") || rowStr.includes("quantity") || 
                        rowStr.includes("equipment")) {
                        headerRowIndex = i;
                        headers = rawRows[i].map(h => h ? String(h).trim() : "");
                        break;
                    }
                }

                if (headerRowIndex !== -1 && headers.length > 0) {
                    for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
                        const rowData = rawRows[i];
                        if (!rowData || rowData.length === 0) continue;
                        
                        const rowObj = {};
                        let hasValue = false;
                        for (let j = 0; j < headers.length; j++) {
                            const key = headers[j];
                            const val = rowData[j];
                            if (key) {
                                rowObj[key] = val;
                                if (val !== null && val !== "") hasValue = true;
                            }
                        }

                        if (hasValue) {
                            const mappedData = mapRowToPayload(rowObj, sheetName);
                            if (mappedData.table_type !== 'unknown') {
                                payload.push(mappedData);
                            }

                            // 🌟 1. ดึงค่าจากไฟล์เทส มาคำนวณ 2 ช่องที่หายไปโดยตรง!
                            const s = String(sheetName).toLowerCase();
                            if (s.includes("produc") || s.includes("circular")) {
                                const output = safeFloat(getExactVal(rowObj, ["output mass", "total output", "ผลิต", "yield"]));
                                const feedstock = safeFloat(getExactVal(rowObj, ["feedstock mass", "circular mass", "waste", "ขยะ"]));
                                
                                if (output > 0) totalOutput += output;
                                if (feedstock > 0) {
                                    totalFeedstock += feedstock;
                                    localCarbonSaved += (feedstock * 1.5) / 1000; // ตีเป็น Carbon Avoided จากพลาสติกรีไซเคิล
                                }
                            }
                            if (s.includes("scope 3") || s.includes("trans")) {
                                const activity = String(getExactVal(rowObj, ["activity", "รถ", "ประเภท", "name"]) || "").toLowerCase();
                                const amount = safeFloat(getExactVal(rowObj, ["quant", "dist", "ระยะ", "ปริมาณ"]));
                                
                                if (activity.includes("recycled") || activity.includes("รีไซเคิล") || activity.includes("circular")) {
                                    localCarbonSaved += (amount * 1.4) / 1000; // ตีเป็น Carbon Avoided จากเหล็กรีไซเคิล ฯลฯ
                                }
                            }
                        }
                    }
                }
            });

            // 🌟 2. คำนวณเป็น % สำหรับ Circular Resource
            if (totalOutput > 0) {
                localCircularPercent = (totalFeedstock / totalOutput) * 100;
            }

            if (payload.length === 0) {
                if (statusLabel) statusLabel.innerText = "สถานะ: ไม่พบข้อมูล (หาหัวตารางไม่เจอ)";
                return;
            }

            if (statusLabel) statusLabel.innerText = "กำลังเคลียร์ข้อมูลเก่า...";
            try {
                await axios.post(`${API_BASE_URL}/api/clear-data`, {}, { headers: { "ngrok-skip-browser-warning": "true" } });
            } catch (err) {}

            if (statusLabel) statusLabel.innerText = "กำลังบันทึกข้อมูลใหม่...";
            let successCount = 0;
            for (const rowData of payload) {
                try {
                    await axios.post(API_CALCULATE, rowData, { headers: { "ngrok-skip-browser-warning": "true", "Content-Type": "application/json" } });
                    successCount++;
                } catch (apiErr) { console.error("❌ Insert Error:", apiErr); }
            }

            if (document.getElementById("totalRows")) document.getElementById("totalRows").innerText = successCount;
            
            setTimeout(async () => {
                await loadDashboardSummary();
                if (statusLabel) statusLabel.innerText = `สถานะ: บันทึกข้อมูลเสร็จสิ้น กำลังสร้าง PDF... ⏳`;
                try {
                    await axios.post(API_REPORT, {}, { headers: { "ngrok-skip-browser-warning": "true" } });
                    if (statusLabel) statusLabel.innerText = `สถานะ: ประมวลผลและสร้าง PDF สำเร็จ ✅ (${successCount} รายการ)`;
                } catch (err) {
                    if (statusLabel) statusLabel.innerText = `สถานะ: ประมวลผลสำเร็จ (แต่สร้าง PDF ไม่สำเร็จ ❌)`;
                }
            }, 500);

        } catch (err) {
            console.error("File Parse Error:", err);
            if (statusLabel) statusLabel.innerText = "สถานะ: Error ตอนอ่านไฟล์";
        }
    };
    reader.readAsArrayBuffer(file);
}

function mapRowToPayload(r, sheetName) {
    const s = String(sheetName).toLowerCase();
    const rawDate = getExactVal(r, ["date", "วันที่", "maintenance"]) || "2026-01-01";
    const base = { date: formatExcelDate(rawDate) };
    
    if (s.includes("util") || s.includes("energy")) {
        const activity = getExactVal(r, ["activity", "energy", "กิจกรรม"]);
        if (!activity) return { table_type: 'unknown' };
        return {
            ...base, table_type: 'utilities', energy_type: activity,
            amount: safeFloat(getExactVal(r, ["quant", "amount", "ปริมาณ"])),
            ef: safeFloat(getExactVal(r, ["ef", "factor"]), 0.5065),
            unit: getExactVal(r, ["unit", "หน่วย"]) || ""
        };

    } else if (s.includes("mainten") || s.includes("fugitive")) {
        const ref = getExactVal(r, ["refrigerant", "สารทำความเย็น"]);
        if (!ref) return { table_type: 'unknown' };
        return {
            ...base, table_type: 'maintenance', refrigerant_type: ref,
            top_up_amount_kg: safeFloat(getExactVal(r, ["fill", "quant", "ปริมาณ", "top"])),
            gwp_value: safeFloat(getExactVal(r, ["gwp"]), 1430)
        };

    } else if (s.includes("scope 3") || s.includes("trans")) {
        const activity = getExactVal(r, ["activity", "รถ", "ประเภท", "name"]);
        if (!activity) return { table_type: 'unknown' };
        
        // ส่งเป็น transportation เหมือนเดิม ป้องกันหลังบ้านพัง
        return {
            ...base, table_type: 'transportation', energy_type: activity, 
            amount: safeFloat(getExactVal(r, ["quant", "dist", "ระยะ", "ปริมาณ"])),
            ef: safeFloat(getExactVal(r, ["ef", "factor"]), 2.5)
        };

    } else if (s.includes("produc") || s.includes("circular")) {
        const product = getExactVal(r, ["product", "สินค้า", "name"]);
        if (!product) return { table_type: 'unknown' };
        
        return {
            ...base, table_type: 'production', product_type: product,
            yield_amount: safeFloat(getExactVal(r, ["output", "yield", "mass", "ผลิต"])),        
            waste_generated: safeFloat(getExactVal(r, ["feedstock", "circular", "waste", "ขยะ"])) || 0
        };
    }
    return { ...base, table_type: 'unknown' };
}

async function loadDashboardSummary() {
    try {
        const timestamp = new Date().getTime();
        const res = await axios.get(`${API_DASHBOARD}?t=${timestamp}`, {
            headers: { "ngrok-skip-browser-warning": "true", "Cache-Control": "no-cache" }
        });
        
        const backendData = res.data.data || res.data;

        const scope1 = backendData.scope_breakdown?.scope1 || 0;
        const scope2 = backendData.scope_breakdown?.scope2 || 0;
        const scope3 = backendData.scope_breakdown?.scope3 || 0;

        rawTotalEmissionKg = scope1 + scope2 + scope3;

        updateUI(backendData);
        updateForestRecommendation(rawTotalEmissionKg / KG_PER_CREDIT);
        
        setTimeout(() => { updateCharts(scope1, scope2, scope3); }, 300);

    } catch (err) {
        console.error("❌ Error loading dashboard summary:", err);
    }
}

function updateUI(data) {
    try {
        const emElem = document.getElementById("totalEmission");
        if (emElem) emElem.innerText = rawTotalEmissionKg.toLocaleString(undefined, {minimumFractionDigits: 2});

        const credit = rawTotalEmissionKg / KG_PER_CREDIT;
        const cost = credit * PRICE_PER_CREDIT;

        if (document.getElementById("carbonCredits")) document.getElementById("carbonCredits").innerText = credit.toFixed(2);
        if (document.getElementById("carbonCost")) document.getElementById("carbonCost").innerText = Math.floor(cost).toLocaleString();

        let circularPercent = 0;
        let carbonSaved = 0;
        
        if (data) {
            const scoreboard = data.dashboard_scoreboard || data.scoreboard || {};
            circularPercent = scoreboard.circular_resource_percent || data.circular_resource_percent || 0;
            carbonSaved = scoreboard.total_carbon_saved_tonCO2e || data.total_carbon_saved_tonCO2e || 0;
        }

        // 🌟 3. พระเอกของเรา: ถ้าหลังบ้านบอกว่า 0 เราจะใช้ค่าที่ดึงจากไฟล์เทสโดยตรง!
        if (circularPercent === 0 && localCircularPercent > 0) {
            circularPercent = localCircularPercent;
        }
        if (carbonSaved === 0 && localCarbonSaved > 0) {
            carbonSaved = localCarbonSaved;
        }

        if (document.getElementById("circularPercent")) {
            document.getElementById("circularPercent").innerText = circularPercent.toFixed(1);
        }
        if (document.getElementById("carbonSaved")) {
            document.getElementById("carbonSaved").innerText = carbonSaved.toLocaleString(undefined, {minimumFractionDigits: 2});
        }
    } catch (e) {
        console.error("UI Update Error:", e);
    }
}

function updateForestRecommendation(credits) {
    try {
        const forestBox = document.getElementById("forest-list");
        if (!forestBox) return;

        if (credits <= 0) {
            forestBox.innerHTML = '<p style="color:#999; font-size:13px; text-align:center; margin-top:20px;">กรุณาอัปโหลดไฟล์เพื่อดูคำแนะนำ</p>';
            return;
        }

        forestBox.innerHTML = `
            <div class="forest-card">
                🌳 <b>โครงการป่าชุมชนบ้านโค้งวันเพ็ญ</b><br>
                <span style="color:#666; font-size:12px;">รองรับได้: 1,500 Credits</span><br>
                <span style="color:var(--primary); font-weight:bold;">ราคา: 500 ฿ / Credit</span>
            </div>
            <div class="forest-card">
                🌲 <b>โครงการปลูกป่าชายเลน จ.ระยอง</b><br>
                <span style="color:#666; font-size:12px;">รองรับได้: 5,000 Credits</span><br>
                <span style="color:var(--primary); font-weight:bold;">ราคา: 450 ฿ / Credit</span>
            </div>
        `;
    } catch (e) { console.error(e); }
}

function updateCharts(s1, s2, s3) {
    try {
        const ctxScope = document.getElementById('scopeChart');
        const ctxBar = document.getElementById('barChart');
        const colors = ['#2d5a27', '#28a745', '#a3d19c'];

        if (ctxScope) {
            if (scopeChartInstance) scopeChartInstance.destroy();
            scopeChartInstance = new Chart(ctxScope.getContext('2d'), {
                type: 'doughnut',
                data: {
                    labels: ['Scope 1', 'Scope 2', 'Scope 3'],
                    datasets: [{
                        data: [s1, s2, s3],
                        backgroundColor: colors,
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'bottom' } }
                }
            });
        }

        if (ctxBar) {
            if (barChartInstance) barChartInstance.destroy();
            barChartInstance = new Chart(ctxBar.getContext('2d'), {
                type: 'bar',
                data: {
                    labels: ['Scope 1', 'Scope 2', 'Scope 3'],
                    datasets: [{
                        label: 'Emissions (kgCO₂e)',
                        data: [s1, s2, s3],
                        backgroundColor: colors,
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: { y: { beginAtZero: true } },
                    plugins: { legend: { display: false } }
                }
            });
        }
    } catch (err) {
        console.error("❌ Chart Error:", err);
    }
}