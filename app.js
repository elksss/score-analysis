// Global State
let studentsDataMid = {};  // Mid-term Exam Data: StudentNum -> Data
let studentsDataFull = {}; // Completed Semesters Data: StudentNum -> Data
let selectedStudentId = null;
let activeTab = 'mid'; // 'mid' or 'full'
let trendChart = null; // Chart.js instance

// Grade scale configurations (5-Grade Scale using rounding rule)
const SCALES = {
    custom: [
        { grade: 1, limit: 0.10 },
        { grade: 2, limit: 0.34 },
        { grade: 3, limit: 0.66 },
        { grade: 4, limit: 0.90 },
        { grade: 5, limit: 1.00 }
    ]
};

// Category Mapping helper
const CATEGORY_MAP = {
    ko: { name: '국어계열', color: '#0c8599', pillClass: 'cat-pill-ko', label: '국어' },
    math: { name: '수학계열', color: '#d6336c', pillClass: 'cat-pill-math', label: '수학' },
    eng: { name: '영어계열', color: '#e08300', pillClass: 'cat-pill-eng', label: '영어' },
    social: { name: '사회계열', color: '#7048e8', pillClass: 'cat-pill-social', label: '사회(역사/도덕포함)' },
    sci: { name: '과학계열', color: '#0ca678', pillClass: 'cat-pill-sci', label: '과학' }
};

// UI Elements
const dropzoneMid = document.getElementById('dropzoneMid');
const btnUploadMid = document.getElementById('btnUploadMid');
const fileInputMid = document.getElementById('fileInputMid');
const fileStatusMid = document.getElementById('fileStatusMid');

const dropzoneFull = document.getElementById('dropzoneFull');
const btnUploadFull = document.getElementById('btnUploadFull');
const fileInputFull = document.getElementById('fileInputFull');
const fileStatusFull = document.getElementById('fileStatusFull');

const controlsSection = document.getElementById('controlsSection');
const studentSelect = document.getElementById('studentSelect');
const tabNavigation = document.getElementById('tabNavigation');
const dashboardContent = document.getElementById('dashboardContent');
const studentNameEl = document.getElementById('studentName');
const studentAvatarText = document.getElementById('studentAvatarText');

// Tab buttons & panes
const tabBtnMid = document.getElementById('tabBtnMid');
const tabBtnFull = document.getElementById('tabBtnFull');
const tabPaneMid = document.getElementById('tabPaneMid');
const tabPaneFull = document.getElementById('tabPaneFull');

// Student Count Badge Element
const studentCountBadge = document.getElementById('studentCountBadge');

// Tab 1 Elements
const subjectGrid = document.getElementById('subjectGrid');
const promptTextMid = document.getElementById('promptTextMid');
const btnCopyPromptMid = document.getElementById('btnCopyPromptMid');
const btnCopyAnalysisMid = document.getElementById('btnCopyAnalysisMid');
const btnCopyAllAnalysisMid = document.getElementById('btnCopyAllAnalysisMid');

// Tab 2 Elements
const semestersTableHeader = document.getElementById('semestersTableHeader');
const semestersTableBody = document.getElementById('semestersTableBody');
const categoryFilters = document.getElementById('categoryFilters');
const promptTextFull = document.getElementById('promptTextFull');
const btnCopyPromptFull = document.getElementById('btnCopyPromptFull');
const btnCopyTableFull = document.getElementById('btnCopyTableFull');

// Event Listeners for File Selection
btnUploadMid.addEventListener('click', () => fileInputMid.click());
fileInputMid.addEventListener('change', () => handleFile(fileInputMid.files[0], 'mid'));

btnUploadFull.addEventListener('click', () => fileInputFull.click());
fileInputFull.addEventListener('change', () => handleFile(fileInputFull.files[0], 'full'));

// Drag & Drop
setupDropzone(dropzoneMid, fileInputMid, 'mid');
setupDropzone(dropzoneFull, fileInputFull, 'full');

function setupDropzone(zone, input, type) {
    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
    });
    zone.addEventListener('dragleave', () => {
        zone.classList.remove('dragover');
    });
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            input.files = e.dataTransfer.files;
            handleFile(input.files[0], type);
        }
    });
}

// Tab Switching
tabBtnMid.addEventListener('click', () => switchTab('mid'));
tabBtnFull.addEventListener('click', () => switchTab('full'));

studentSelect.addEventListener('change', (e) => {
    selectedStudentId = e.target.value;
    updateTabsState();
    renderDashboard();
});

btnCopyPromptMid.addEventListener('click', () => copyToClipboard(promptTextMid, btnCopyPromptMid));
btnCopyPromptFull.addEventListener('click', () => copyToClipboard(promptTextFull, btnCopyPromptFull));
if (btnCopyAnalysisMid) {
    btnCopyAnalysisMid.addEventListener('click', copyAnalysisMidToClipboard);
}
if (btnCopyAllAnalysisMid) {
    btnCopyAllAnalysisMid.addEventListener('click', copyAllAnalysisMidToClipboard);
}
if (btnCopyTableFull) {
    btnCopyTableFull.addEventListener('click', copyTableToClipboard);
}

// File Handling
function handleFile(file, type) {
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Auto-detect type based on file contents
            const isFullFile = checkIsFullFile(sheetData);
            
            if (isFullFile) {
                parseFullFile(sheetData);
                fileStatusFull.textContent = file.name;
                fileStatusFull.classList.add('loaded');
            } else {
                parseMidFile(sheetData);
                fileStatusMid.textContent = file.name;
                fileStatusMid.classList.add('loaded');
            }

            populateStudentSelector();
            
            // Show student selector controls
            controlsSection.classList.remove('hidden');
        } catch (error) {
            console.error(error);
            alert('엑셀 파일을 읽는 도중 오류가 발생했습니다. 올바른 형식인지 확인하세요.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Checks if file is the Completed Semesters file based on report card header cell
function checkIsFullFile(grid) {
    for (let r = 0; r < Math.min(10, grid.length); r++) {
        const row = grid[r];
        if (row) {
            for (let c = 0; c < row.length; c++) {
                const val = String(row[c]);
                if (val.includes('학교생활기록부') || val.includes('교과학습발달상황')) {
                    return true;
                }
            }
        }
    }
    return false;
}

// Tab Switching Handler
function switchTab(tab) {
    if (tab === 'mid' && tabBtnMid.classList.contains('disabled')) return;
    if (tab === 'full' && tabBtnFull.classList.contains('disabled')) return;

    activeTab = tab;
    if (tab === 'mid') {
        tabBtnMid.classList.add('active');
        tabBtnFull.classList.remove('active');
        tabPaneMid.classList.add('active');
        tabPaneFull.classList.remove('active');
    } else {
        tabBtnMid.classList.remove('active');
        tabBtnFull.classList.add('active');
        tabPaneMid.classList.remove('active');
        tabPaneFull.classList.add('active');
        // Redraw chart and table when Tab 2 opens to fix sizing issues
        setTimeout(renderTabFull, 50);
    }
}

// Enable/Disable tabs based on uploaded files for the selected student
function updateTabsState() {
    const hasMid = studentsDataMid[selectedStudentId] !== undefined;
    const hasFull = studentsDataFull[selectedStudentId] !== undefined;

    if (hasMid) {
        tabBtnMid.classList.remove('disabled');
    } else {
        tabBtnMid.classList.add('disabled');
    }

    if (hasFull) {
        tabBtnFull.classList.remove('disabled');
    } else {
        tabBtnFull.classList.add('disabled');
    }

    // Auto-switch tab if the active one is disabled
    if (activeTab === 'mid' && !hasMid && hasFull) {
        switchTab('full');
    } else if (activeTab === 'full' && !hasFull && hasMid) {
        switchTab('mid');
    } else {
        switchTab(activeTab);
    }

    if (hasMid || hasFull) {
        tabNavigation.classList.remove('hidden');
    } else {
        tabNavigation.classList.add('hidden');
    }
}

// Populate Student Selector Dropdown (Union of both datasets)
function populateStudentSelector() {
    const allIds = new Set([
        ...Object.keys(studentsDataMid),
        ...Object.keys(studentsDataFull)
    ]);
    
    const sortedIds = Array.from(allIds).sort((a, b) => parseInt(a) - parseInt(b));
    
    studentSelect.innerHTML = '<option value="" disabled selected>학생을 선택하세요</option>';
    
    sortedIds.forEach(id => {
        const name = studentsDataMid[id]?.name || studentsDataFull[id]?.name || "알수없음";
        const option = document.createElement('option');
        option.value = id;
        option.textContent = `${id}번 ${name}`;
        studentSelect.appendChild(option);
    });

    if (studentCountBadge) {
        studentCountBadge.textContent = `로딩 완료: ${sortedIds.length}명`;
        studentCountBadge.classList.remove('hidden');
    }

    if (selectedStudentId) {
        studentSelect.value = selectedStudentId;
    }
}

// Clipboard copy utility
function copyToClipboard(textarea, btn) {
    textarea.select();
    navigator.clipboard.writeText(textarea.value).then(() => {
        const originalText = btn.innerHTML;
        btn.innerHTML = '<i class="fa-solid fa-check"></i> 복사 완료';
        btn.style.backgroundColor = '#0ca678';
        setTimeout(() => {
            btn.innerHTML = originalText;
            btn.style.backgroundColor = '';
        }, 1500);
    });
}

// ==========================================
// PARSING LOGIC: Tab 1 (1회고사 성적 파일)
// ==========================================
function parseMidFile(grid) {
    studentsDataMid = {};
    const labelWonjeomsu = "원점수";
    const labelSeokcha = "석차";
    const labelSugangjasu = "수강자수";

    const tableBlocks = [
        { subRow: 4, labelRow: 6, startRow: 7, endRow: 24 }, // Table 1
        { subRow: 26, labelRow: 28, startRow: 29, endRow: 46 }, // Table 2
        { subRow: 48, labelRow: 50, startRow: 51, endRow: 68 }  // Table 3 (completed values start)
    ];

    const colLetters = [];
    for (let i = 0; i < 100; i++) colLetters.push(colIndexToLetter(i));

    tableBlocks.forEach(block => {
        // Adjust for index alignment
        const subRow = grid[block.subRow - 1];
        const labelRow = grid[block.labelRow - 1];
        if (!subRow || !labelRow) return;

        const subjects = [];
        let currentSub = null;

        for (let cIdx = 2; cIdx < 80; cIdx++) {
            const col = colLetters[cIdx];
            const cellVal = subRow[cIdx];

            if (cellVal && typeof cellVal === 'string' && cellVal.includes('(')) {
                if (currentSub) {
                    currentSub.endCol = colLetters[cIdx - 1];
                    subjects.push(currentSub);
                }
                currentSub = {
                    name: cellVal.trim(),
                    startCol: col,
                    endCol: col,
                    rawCol: null,
                    rankCol: null,
                    countCol: null
                };
            }
        }
        if (currentSub) {
            currentSub.endCol = colLetters[79];
            subjects.push(currentSub);
        }

        subjects.forEach(sub => {
            const startIdx = colLetterToIndex(sub.startCol);
            const endIdx = colLetterToIndex(sub.endCol);

            for (let i = startIdx; i <= endIdx; i++) {
                const val = labelRow[i];
                if (val && typeof val === 'string') {
                    const cleanVal = val.replace(/\r?\n/g, "").trim();
                    if (cleanVal === labelWonjeomsu) {
                        sub.rawCol = i;
                    } else if (cleanVal.includes(labelSeokcha)) {
                        sub.rankCol = i;
                    } else if (cleanVal === labelSugangjasu) {
                        sub.countCol = i;
                    }
                }
            }
        });

        for (let r = block.startRow - 1; r <= block.endRow - 1; r++) {
            const row = grid[r];
            if (!row) continue;

            const numVal = row[0];
            const nameVal = row[1];
            if (!nameVal) continue;

            const name = String(nameVal).trim();
            const num = String(numVal).trim();
            const key = num;

            if (!studentsDataMid[key]) {
                studentsDataMid[key] = {
                    id: key,
                    num: num,
                    name: name,
                    subjects: {}
                };
            }

            subjects.forEach(sub => {
                if (sub.rawCol !== null && sub.rankCol !== null && sub.countCol !== null) {
                    const rawVal = row[sub.rawCol];
                    const rankVal = row[sub.rankCol];
                    const countVal = row[sub.countCol];

                    if (rawVal !== undefined || rankVal !== undefined) {
                        const parsedScore = rawVal !== undefined ? parseFloat(rawVal) : null;
                        const parsedCount = countVal !== undefined ? parseInt(countVal) : null;

                        studentsDataMid[key].subjects[sub.name] = {
                            score: isNaN(parsedScore) ? rawVal : parsedScore,
                            rankRaw: rankVal !== undefined ? String(rankVal).trim() : null,
                            totalStudents: isNaN(parsedCount) ? null : parsedCount
                        };
                    }
                }
            });
        }
    });
}

// ==========================================
// PARSING LOGIC: Tab 2 (전학기 성적 파일)
// ==========================================
function parseFullFile(grid) {
    studentsDataFull = {};
    
    let activeNum = null;
    let activeName = null;
    let activeGrade = null;
    let activeSemester = null;

    for (let r = 4; r < grid.length; r++) {
        const row = grid[r];
        if (!row) continue;

        const colA = row[0];
        if (colA && String(colA).includes('이수학점')) continue; // Skip totals

        // Check if student starts (indicated by numeric student number in col A)
        if (colA && /^\d+$/.test(String(colA).trim())) {
            activeNum = String(colA).trim();
            activeName = String(row[1]).trim();
            activeGrade = String(row[2]).trim();
            activeSemester = String(row[3]).trim();
        } else if (!colA && row[3]) {
            // Carriage return within same student, check for semester change
            const semVal = String(row[3]).trim();
            if (/^\d+$/.test(semVal)) {
                activeSemester = semVal;
            }
        }

        // Parse subject record (Col E must be a category, check against E index)
        const colE = row[4]; // 교과
        const colF = row[5]; // 과목
        if (colE && colF) {
            const cleanColE = String(colE).replace(/\s+/g, '');
            const cleanColF = String(colF).replace(/\s+/g, '');
            if (cleanColE === "교과" || cleanColF === "과목") continue;

            const credits = parseInt(row[6]); // 학점 (단위수)
            if (isNaN(credits)) continue;

            const category = String(colE).trim();
            const subjectName = String(colF).trim();
            
            const scoreText = row[7] !== undefined ? String(row[7]).trim() : "";
            const rawScore = scoreText.includes('/') ? scoreText.split('/')[0] : scoreText;
            
            const achievement = row[8] !== undefined ? String(row[8]).trim() : "";
            const grade = row[10] !== undefined ? parseInt(row[10]) : NaN;
            const totalStudents = row[11] !== undefined ? parseInt(row[11]) : null;

            if (activeNum && activeName) {
                const key = activeNum;
                if (!studentsDataFull[key]) {
                    studentsDataFull[key] = {
                        id: key,
                        num: activeNum,
                        name: activeName,
                        semesters: {}
                    };
                }

                const semKey = `${activeGrade}학년 ${activeSemester}학기`;
                if (!studentsDataFull[key].semesters[semKey]) {
                    studentsDataFull[key].semesters[semKey] = [];
                }

                studentsDataFull[key].semesters[semKey].push({
                    category: category,
                    name: subjectName,
                    credits: credits,
                    rawScore: rawScore,
                    achievement: achievement,
                    grade: grade,
                    totalStudents: totalStudents
                });
            }
        }
    }
    console.log("Parsed Completed Semesters:", studentsDataFull);
}

// Calculate Grade based on rank, ties, total students, and current scale (using Rounding rule)
function calculateGrade(rank, ties, totalStudents) {
    const intermediateRank = rank + (ties - 1) / 2;
    const scale = SCALES.custom;
    for (let i = 0; i < scale.length; i++) {
        // 등급별 누적 인원수 = Math.round(수강자수 * 누적비율)
        const allowedCount = Math.round(totalStudents * scale[i].limit);
        if (intermediateRank <= allowedCount) {
            return scale[i].grade;
        }
    }
    return scale[scale.length - 1].grade; // Return worst grade
}

// Get the cutoff percentage for a specific target grade
function getTargetLimit(targetGrade) {
    const scale = SCALES.custom;
    const match = scale.find(s => s.grade === targetGrade);
    return match ? match.limit : null;
}

// ==========================================
// RENDER LOGIC
// ==========================================
function renderDashboard() {
    dashboardContent.classList.remove('hidden');

    const hasMid = studentsDataMid[selectedStudentId] !== undefined;
    const hasFull = studentsDataFull[selectedStudentId] !== undefined;
    const name = studentsDataMid[selectedStudentId]?.name || studentsDataFull[selectedStudentId]?.name || "알수없음";

    studentNameEl.innerHTML = `${name} <span class="student-number" id="studentNum">2학년 2반 ${selectedStudentId}번</span>`;
    studentAvatarText.textContent = name.charAt(name.length - 1);

    if (hasMid) renderTabMid();
    if (hasFull) renderTabFull();
}

// Render Tab 1 (1회고사 분석)
function renderTabMid() {
    const student = studentsDataMid[selectedStudentId];
    if (!student) return;

    const subjectList = [];
    let sumWeightedGrades = 0;
    let sumUnits = 0;

    Object.keys(student.subjects).forEach(subName => {
        const subData = student.subjects[subName];
        
        // Extract Unit Weight e.g. "대수(4)" -> 4
        const unitMatch = subName.match(/\((\d+)\)/);
        const units = unitMatch ? parseInt(unitMatch[1]) : 3;

        const { rank, ties } = parseRank(subData.rankRaw);
        if (rank === null || subData.totalStudents === null) return;

        const intermediateRank = rank + (ties - 1) / 2;
        const percentage = intermediateRank / subData.totalStudents;

        // Current grade using rounding rule
        const currentGrade = calculateGrade(rank, ties, subData.totalStudents);

        const targetGrade = Math.max(1, currentGrade - 1);
        const targetLimit = getTargetLimit(targetGrade);

        let targetRank = null;
        let overtake = 0;

        if (targetLimit !== null && currentGrade > 1) {
            targetRank = Math.round(subData.totalStudents * targetLimit);
            overtake = Math.max(0, rank - targetRank);
        }

        sumWeightedGrades += currentGrade * units;
        sumUnits += units;

        subjectList.push({
            name: subName,
            units: units,
            score: subData.score,
            rankRaw: subData.rankRaw,
            rank: rank,
            totalStudents: subData.totalStudents,
            percentage: percentage,
            currentGrade: currentGrade,
            targetGrade: targetGrade,
            targetRank: targetRank,
            overtake: overtake
        });
    });

    // Sort subjects by overtake count (ascending)
    subjectList.sort((a, b) => {
        if (a.overtake !== b.overtake) {
            return a.overtake - b.overtake;
        }
        return b.percentage - a.percentage;
    });

    // Render Cards
    subjectGrid.innerHTML = "";
    subjectList.forEach((sub, idx) => {
        const cardIdx = idx + 1;
        const card = document.createElement('div');
        card.className = 'subject-card';
        
        if (cardIdx === 1) card.classList.add('card-priority-1');
        else if (cardIdx === 2) card.classList.add('card-priority-2');
        else if (cardIdx === 3) card.classList.add('card-priority-3');
        else card.classList.add('card-priority-other');

        let badgeHtml = "";
        if (cardIdx === 1) badgeHtml = '<i class="fa-solid fa-medal medal-gold"></i>';
        else if (cardIdx === 2) badgeHtml = '<i class="fa-solid fa-medal medal-silver"></i>';
        else if (cardIdx === 3) badgeHtml = '<i class="fa-solid fa-medal medal-bronze"></i>';
        else badgeHtml = `<span class="rank-text-badge">${cardIdx}위</span>`;

        let overtakeMessage = "";
        if (sub.currentGrade === 1) {
            overtakeMessage = '<span class="overtake-message" style="color: #0ca678;">최상위 등급 유지 중!</span>';
        } else {
            overtakeMessage = `앞에 <span class="overtake-count" style="font-weight: 700; font-size: 1.05rem;">${sub.overtake}명</span> 추월 필요`;
        }

        let targetRankText = "";
        if (sub.currentGrade === 1) {
            targetRankText = `목표 1위 유지`;
        } else {
            targetRankText = `목표 ${sub.targetRank}위 이내`;
        }

        const cleanName = sub.name.replace(/\(\d+\)/, "").trim();

        card.innerHTML = `
            <div class="priority-badge-container">
                ${badgeHtml}
            </div>
            <div class="card-content">
                <div class="subject-title-row">
                    <span class="subject-name">${cleanName}(${sub.units})</span>
                    <span class="grade-transition-badge">
                        <span class="grade-current">${sub.currentGrade}등급</span>
                        ${sub.currentGrade > 1 ? `<i class="fa-solid fa-arrow-right"></i><span class="grade-target">${sub.targetGrade}등급</span>` : ""}
                    </span>
                </div>
                <div class="score-details-row">
                    <span><i class="fa-solid fa-location-pin icon-pin"></i> 현재 ${sub.rankRaw}위 / ${sub.totalStudents}명</span>
                    <span class="raw-score-indicator">원점수 ${sub.score}점</span>
                </div>
                <div class="target-summary-row">
                    <span class="target-badge"><i class="fa-solid fa-bullseye icon-bullseye"></i> ${targetRankText}</span>
                    <span class="overtake-message">${overtakeMessage}</span>
                </div>
            </div>
        `;
        subjectGrid.appendChild(card);
    });

    // Populate Tab 1 AI Prompt
    let subjectsPromptText = "";
    subjectList.forEach((sub, idx) => {
        const cleanName = sub.name.replace(/\(\d+\)/, "").trim();
        subjectsPromptText += `${idx + 1}. ${cleanName}(${sub.units}단위): 현재 ${sub.currentGrade}등급(원점수 ${sub.score}점, ${sub.rankRaw}위/${sub.totalStudents}명) -> 목표 ${sub.targetGrade}등급(목표 ${sub.targetRank !== null ? sub.targetRank + '위 이내' : '유지'}, 앞에 ${sub.overtake}명 추월 필요)\n`;
    });

    const avgGrade = sumUnits > 0 ? (sumWeightedGrades / sumUnits).toFixed(2) : "-";
    const avgGradeMidEl = document.getElementById('avgGradeMid');
    if (avgGradeMidEl) {
        avgGradeMidEl.textContent = avgGrade !== "-" ? `${avgGrade}등급` : "-";
    }
    promptTextMid.value = `[학생 성적 분석 및 2회고사 대비 학습 전략 제안]
너는 고등학교 성적 상담 전문 AI 컨설턴트야. 아래 제공된 학생의 1학기 1회고사 성적 데이터를 바탕으로, 2회고사 성적 향상을 위한 구체적이고 실용적인 맞춤형 학습 전략을 작성해줘. 추상적인 조언은 피하고, 과목별 실제 수치(점수, 석차, 추월 인원)를 근거로 행동 방침을 제안해줘.

[학생 기본 정보]
- 학생 이름: ${student.name} (${student.num}번)
- 평균 등급: ${avgGrade}등급 (1회고사 기준)
- 분석 대상 과목 수: ${subjectList.length}개

[과목별 상세 데이터 (우선순위 순 - 등급 올리기 가장 쉬운 순 정렬)]
${subjectsPromptText}
[요청 사항]
1. 가장 빠르게 등급 향상이 가능한 '우선 공략 과목'(예: 추월 인원이 0~2명으로 조금만 노력하면 등급이 오르는 과목)을 선정하고, 왜 이 과목을 먼저 공부해야 하는지 수치적으로 설명해줘.
2. 현재 등급이 우수하여 유지해야 하는 과목과, 등급 경계선에 아슬아슬하게 걸쳐 있어 2회고사 때 방어 및 추월이 필요한 과목을 구분해줘.
3. 각 과목의 '단위수'를 고려하여 총합 등급 평점(GPA)을 극대화하기 위해 다음 2회고사를 준비하는 과목별 시간 배분(공부 비율%)을 구체적으로 제안해줘.`;
}

function renderTabFull() {
    const student = studentsDataFull[selectedStudentId];
    if (!student) return;

    // 1. Find all chronological semesters
    const semesters = Object.keys(student.semesters).sort((a, b) => {
        const parse = (str) => {
            const m = str.match(/(\d+)학년\s+(\d+)학기/);
            return m ? parseInt(m[1]) * 10 + parseInt(m[2]) : 0;
        };
        return parse(a) - parse(b);
    });

    // Compute GPAs (Semester, Yearly, Cumulative)
    let cumulativeSumWeighted = 0;
    let cumulativeSumCredits = 0;
    const semesterGPAs = {}; 
    const yearlyGPAs = {}; 

    semesters.forEach(sem => {
        let semSumWeighted = 0;
        let semSumCredits = 0;
        const semList = student.semesters[sem] || [];

        semList.forEach(sub => {
            if (!isNaN(sub.grade)) {
                semSumWeighted += sub.grade * sub.credits;
                semSumCredits += sub.credits;
                
                cumulativeSumWeighted += sub.grade * sub.credits;
                cumulativeSumCredits += sub.credits;

                const yearMatch = sem.match(/(\d+)학년/);
                if (yearMatch) {
                    const yearKey = yearMatch[1] + "학년";
                    if (!yearlyGPAs[yearKey]) {
                        yearlyGPAs[yearKey] = { sumWeighted: 0, sumCredits: 0 };
                    }
                    yearlyGPAs[yearKey].sumWeighted += sub.grade * sub.credits;
                    yearlyGPAs[yearKey].sumCredits += sub.credits;
                }
            }
        });

        if (semSumCredits > 0) {
            semesterGPAs[sem] = (semSumWeighted / semSumCredits).toFixed(2);
        } else {
            semesterGPAs[sem] = null;
        }
    });

    const cumulativeGPA = cumulativeSumCredits > 0 ? (cumulativeSumWeighted / cumulativeSumCredits).toFixed(2) : "-";

    const cumulativeAvgEl = document.getElementById('cumulativeAvgEl');
    if (cumulativeAvgEl) {
        cumulativeAvgEl.textContent = cumulativeGPA !== "-" ? `${cumulativeGPA}등급` : "-";
    }

    const averagesBreakdownList = document.getElementById('averagesBreakdownList');
    if (averagesBreakdownList) {
        averagesBreakdownList.innerHTML = "";
        const uniqueYears = Object.keys(yearlyGPAs).sort((a, b) => parseInt(a) - parseInt(b));
        
        uniqueYears.forEach(year => {
            const yData = yearlyGPAs[year];
            const yAvg = yData.sumCredits > 0 ? (yData.sumWeighted / yData.sumCredits).toFixed(2) : "-";
            
            const yearItem = document.createElement('div');
            yearItem.className = 'breakdown-item-year';
            yearItem.innerHTML = `${year} 전체 평균: <span style="color: var(--primary-color);">${yAvg}등급</span>`;
            averagesBreakdownList.appendChild(yearItem);

            semesters.forEach(sem => {
                if (sem.startsWith(year) && semesterGPAs[sem] !== null) {
                    const semItem = document.createElement('div');
                    semItem.className = 'breakdown-item-sem';
                    semItem.innerHTML = `${sem}: <span style="font-weight: 600;">${semesterGPAs[sem]}등급</span>`;
                    averagesBreakdownList.appendChild(semItem);
                }
            });
        });
    }

    // Populate semester column headers
    semestersTableHeader.innerHTML = '<th>과목</th>';
    semesters.forEach(sem => {
        const th = document.createElement('th');
        th.textContent = sem;
        semestersTableHeader.appendChild(th);
    });

    // 2. Identify all unique subjects taken by this student across all semesters
    const uniqueSubjects = [];
    semesters.forEach(sem => {
        student.semesters[sem]?.forEach(sub => {
            const exists = uniqueSubjects.find(s => s.name === sub.name);
            if (!exists) {
                uniqueSubjects.push({
                    name: sub.name,
                    credits: sub.credits,
                    category: sub.category
                });
            }
        });
    });

    // Sort subjects by category order
    const catOrder = ['국어', '수학', '영어', '사회(역사/도덕포함)', '과학'];
    uniqueSubjects.sort((a, b) => {
        let idxA = catOrder.indexOf(a.category);
        let idxB = catOrder.indexOf(b.category);
        if (idxA === -1) idxA = 99;
        if (idxB === -1) idxB = 99;
        if (idxA !== idxB) return idxA - idxB;
        return a.name.localeCompare(b.name, 'ko');
    });

    // Populate table rows
    semestersTableBody.innerHTML = "";
    uniqueSubjects.forEach(sub => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td style="font-weight: 600; text-align: left; padding-left: 15px;">${sub.name}(${sub.credits})</td>`;

        semesters.forEach(sem => {
            const matched = student.semesters[sem]?.find(s => s.name === sub.name);
            if (matched) {
                if (!isNaN(matched.grade)) {
                    tr.innerHTML += `
                        <td>
                            <span class="table-grade-badge table-grade-${matched.grade}">${matched.grade}등급</span>
                            <span class="table-score">${matched.rawScore}점 (${matched.achievement})</span>
                            <span class="table-count">${matched.totalStudents ? matched.totalStudents + '명' : '-'}</span>
                        </td>
                    `;
                } else {
                    const gradeClass = matched.rawScore === 'P' ? 'p' : 'none';
                    tr.innerHTML += `
                        <td>
                            <span class="table-grade-badge table-grade-${gradeClass}">${matched.rawScore}</span>
                            <span class="table-score">${matched.achievement ? '성취도: ' + matched.achievement : '-'}</span>
                            <span class="table-count">${matched.totalStudents ? matched.totalStudents + '명' : '-'}</span>
                        </td>
                    `;
                }
            } else {
                tr.innerHTML += '<td>-</td>';
            }
        });
        semestersTableBody.appendChild(tr);
    });

    // 3. Calculate Semester Averages & Categories counts for Trend Chart
    const categories = ['국어', '수학', '영어', '사회(역사/도덕포함)', '과학'];
    const categoryTrendData = {}; // Cat -> Array of grades (null if no data)
    categories.forEach(cat => { categoryTrendData[cat] = []; });
    const allSubjectGrades = []; // Semester averages

    semesters.forEach(sem => {
        let sumWeighted = 0;
        let sumCredits = 0;
        const semList = student.semesters[sem] || [];

        categories.forEach(cat => {
            // Find all subjects in this category with a valid grade
            const matchedSubjects = semList.filter(s => s.category === cat && !isNaN(s.grade));
            if (matchedSubjects.length > 0) {
                let catSumWeighted = 0;
                let catSumCredits = 0;
                matchedSubjects.forEach(s => {
                    catSumWeighted += s.grade * s.credits;
                    catSumCredits += s.credits;
                });
                const catAvg = catSumCredits > 0 ? catSumWeighted / catSumCredits : null;
                categoryTrendData[cat].push(catAvg ? parseFloat(catAvg.toFixed(2)) : null);
                
                sumWeighted += catSumWeighted;
                sumCredits += catSumCredits;
            } else {
                categoryTrendData[cat].push(null);
            }
        });

        if (sumCredits > 0) {
            allSubjectGrades.push(parseFloat((sumWeighted / sumCredits).toFixed(2)));
        } else {
            allSubjectGrades.push(null);
        }
    });

    // Render Pills
    categoryFilters.innerHTML = "";
    
    // Add "전체과목" Pill
    const pillAll = document.createElement('button');
    pillAll.className = 'cat-pill cat-pill-all active';
    pillAll.innerHTML = `<i class="fa-solid fa-list"></i> 전체과목 (${categories.length})`;
    pillAll.addEventListener('click', () => filterChart('all'));
    categoryFilters.appendChild(pillAll);

    // Add specific Category Pills
    Object.keys(CATEGORY_MAP).forEach(key => {
        const catMap = CATEGORY_MAP[key];
        const pill = document.createElement('button');
        pill.className = `cat-pill ${catMap.pillClass}`;
        
        // Count subjects in this category
        let count = 0;
        semesters.forEach(sem => {
            const matchedCount = student.semesters[sem]?.filter(s => s.category === catMap.label).length || 0;
            if (matchedCount > count) count = matchedCount;
        });

        pill.innerHTML = `${catMap.label}계열 (${count})`;
        pill.addEventListener('click', () => filterChart(key));
        categoryFilters.appendChild(pill);
    });

    // Render Line Chart
    renderTrendChart(semesters, categoryTrendData, allSubjectGrades);

    // Populate Tab 2 AI Prompt
    let trendPromptText = "";
    semesters.forEach((sem, sIdx) => {
        trendPromptText += `- ${sem}: 전체평균 ${allSubjectGrades[sIdx] || "-"}등급\n`;
        const semDetails = [];
        student.semesters[sem]?.forEach(sub => {
            if (!isNaN(sub.grade)) {
                semDetails.push(`  * ${sub.name}(${sub.credits}단위): ${sub.grade}등급 (원점수 ${sub.rawScore}점, 성취도 ${sub.achievement}, 수강자수 ${sub.totalStudents}명)`);
            } else {
                semDetails.push(`  * ${sub.name}(${sub.credits}단위): 성취도 ${sub.achievement || '-'} (원점수/이수여부 ${sub.rawScore})`);
            }
        });
        trendPromptText += semDetails.join("\n") + "\n";
    });

    let totalEarnedCredits = 0;
    semesters.forEach(sem => {
        student.semesters[sem]?.forEach(sub => { totalEarnedCredits += sub.credits; });
    });

    promptTextFull.value = `[전학기 성적 추이 분석 및 다음 학기 대비 전략 제안]
너는 고등학교 성적 상담 전문 AI 컨설턴트야. 아래 제공된 학생의 전학기 누적 성적 추이 데이터를 바탕으로, 다음 학기 대비를 위한 구체적이고 실용적인 진로 및 학습 로드맵을 작성해줘. 실제 성적 수치와 계열별 등급 변화 흐름을 근거로 분석해줘.

[학생 기본 정보]
- 학생 이름: ${student.name} (${student.num}번)
- 누적 이수 단위수: ${totalEarnedCredits}단위

[학기별 성적 추이 및 과목 점수 이력]
${trendPromptText}
[요청 사항]
1. 국어, 수학, 영어, 사회, 과학 계열별 등급 변화 추이 중 가장 성취도가 높은 계열과 하락/상승세가 뚜렷한 계열을 데이터 분석을 통해 파악하고 설명해줘.
2. 이 학생의 강점 계열(인문사회 vs 자연공학)을 성적 수치(단위수 곱한 평균 등급)를 근거로 판별하고, 학생부 종합전형 대비를 위한 다음 학기 집중 공략 과목군을 추천해줘.
3. 특정 학기 또는 특정 과목에서 성적이 급락했거나 약점이 드러난 부분을 짚어내고, 다음 학기 시작 전 방학 동안 이를 보완하기 위한 실천 전략을 제시해줘.`;
}

// Chart.js Category Line Chart Renderer
function renderTrendChart(semesters = [], trendData = {}, avgGrades = []) {
    const ctx = document.getElementById('trendChart');
    if (!ctx) return;

    if (trendChart) {
        trendChart.destroy();
    }

    // Default label: all high school semesters (inverted to 6 semesters)
    const labels = ['1학년 1학기', '1학년 2학기', '2학년 1학기', '2학년 2학기', '3학년 1학기', '3학년 2학기'];
    
    // Create datasets
    const datasets = [];

    // Overall Average GPA (thick dashed line)
    const avgData = Array(6).fill(null);
    semesters.forEach((sem, idx) => {
        const lIdx = labels.indexOf(sem);
        if (lIdx !== -1) avgData[lIdx] = avgGrades[idx];
    });

    datasets.push({
        label: '전체과목(평균)',
        data: avgData,
        borderColor: '#495057',
        borderWidth: 3,
        borderDash: [5, 5],
        backgroundColor: 'transparent',
        tension: 0.15,
        segment: {
            borderColor: ctx => (ctx.p0 && ctx.p0.skip) || (ctx.p1 && ctx.p1.skip) ? undefined : '#495057',
        },
        spanGaps: true,
        key: 'all'
    });

    // Individual category lines
    Object.keys(CATEGORY_MAP).forEach(key => {
        const catMap = CATEGORY_MAP[key];
        const catData = Array(6).fill(null);
        
        semesters.forEach((sem, idx) => {
            const lIdx = labels.indexOf(sem);
            if (lIdx !== -1) {
                const gradeList = trendData[catMap.label];
                catData[lIdx] = gradeList ? gradeList[idx] : null;
            }
        });

        datasets.push({
            label: catMap.label,
            data: catData,
            borderColor: catMap.color,
            backgroundColor: catMap.color,
            borderWidth: 2,
            pointRadius: 4,
            pointHoverRadius: 6,
            tension: 0.15,
            spanGaps: true,
            key: key
        });
    });

    trendChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    reverse: true, // 1등급이 맨 위에 오도록 설정!
                    min: 0.7,
                    max: 5.3,
                    afterBuildTicks: function(scale) {
                        scale.ticks = [1, 2, 3, 4, 5].map(v => ({ value: v }));
                    },
                    ticks: {
                        callback: function(value) {
                            return value + '등급';
                        }
                    },
                    grid: {
                        color: '#f1f3f5'
                    }
                },
                x: {
                    grid: {
                        display: false
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        boxWidth: 12,
                        padding: 15,
                        font: {
                            family: 'Noto Sans KR'
                        }
                    }
                }
            }
        }
    });
}

// Chart Category Filter (Click Pills)
function filterChart(key) {
    // Update active pill styling
    const pills = categoryFilters.querySelectorAll('.cat-pill');
    pills.forEach(p => p.classList.remove('active'));

    const activePill = categoryFilters.querySelector(`.cat-pill-${key}`) || categoryFilters.querySelector('.cat-pill-all');
    if (activePill) activePill.classList.add('active');

    if (!trendChart) return;

    trendChart.data.datasets.forEach(dataset => {
        if (key === 'all') {
            // Show all lines
            dataset.hidden = false;
        } else {
            // Show only matched category line or overall avg
            if (dataset.key === key || dataset.key === 'all') {
                dataset.hidden = false;
            } else {
                dataset.hidden = true;
            }
        }
    });
    trendChart.update();
}

// Helper: Convert column letter to 0-based index
function colLetterToIndex(col) {
    let idx = 0;
    for (let i = 0; i < col.length; i++) {
        idx = idx * 26 + (col.charCodeAt(i) - 65 + 1); // 65 is 'A'
    }
    return idx - 1;
}

// Helper: Convert 0-based index to column letter
function colIndexToLetter(idx) {
    let col = "";
    let temp = idx + 1;
    while (temp > 0) {
        let mod = (temp - 1) % 26;
        col = String.fromCharCode(65 + mod) + col;
        temp = Math.floor((temp - mod) / 26);
    }
    return col;
}

// Helper: Parse raw rank string (e.g., "1" or "1 (2)" or "3(4)") into rank and ties
function parseRank(rankRaw) {
    if (rankRaw === undefined || rankRaw === null) return { rank: null, ties: 1 };
    
    // Convert to string and clean
    const str = String(rankRaw).trim();
    if (!str) return { rank: null, ties: 1 };
    
    // Check if there is a parenthesis indicating ties, e.g. "1 (2)" or "2(3)" or "1(2)"
    const match = str.match(/^(\d+)(?:\s*\(\s*(\d+)\s*\))?/);
    if (match) {
        const rank = parseInt(match[1], 10);
        const ties = match[2] ? parseInt(match[2], 10) : 1;
        return { rank, ties };
    }
    
    const val = parseInt(str, 10);
    return { rank: isNaN(val) ? null : val, ties: 1 };
}

// Copy Tab 1 Analysis Results to Clipboard
function copyAnalysisMidToClipboard() {
    const student = studentsDataMid[selectedStudentId];
    if (!student) return;
    
    const subjectList = [];
    Object.keys(student.subjects).forEach(subName => {
        const subData = student.subjects[subName];
        const unitMatch = subName.match(/\((\d+)\)/);
        const units = unitMatch ? parseInt(unitMatch[1]) : 3;
        const { rank, ties } = parseRank(subData.rankRaw);
        if (rank === null || subData.totalStudents === null) return;

        const currentGrade = calculateGrade(rank, ties, subData.totalStudents);
        const targetGrade = Math.max(1, currentGrade - 1);
        const targetLimit = getTargetLimit(targetGrade);
        let targetRank = null;
        let overtake = 0;
        if (targetLimit !== null && currentGrade > 1) {
            targetRank = Math.round(subData.totalStudents * targetLimit);
            overtake = Math.max(0, rank - targetRank);
        }

        subjectList.push({
            name: subName.replace(/\(\d+\)/, "").trim(),
            units: units,
            score: subData.score,
            rankRaw: subData.rankRaw,
            totalStudents: subData.totalStudents,
            currentGrade: currentGrade,
            targetGrade: targetGrade,
            targetRank: targetRank,
            overtake: overtake
        });
    });

    subjectList.sort((a, b) => a.overtake - b.overtake);

    const lines = [
        `[1회고사 성적 분석 결과 - ${student.name} (${selectedStudentId}번)]`,
        `평균 등급: ${calculateAverageGPA(student)}등급`,
        `과목별 2회고사 목표 및 추월 필요 인원 (우선순위 순):`
    ];

    subjectList.forEach((sub, idx) => {
        if (sub.currentGrade === 1) {
            lines.push(`${idx + 1}. ${sub.name}(${sub.units}단위): 현재 1등급 (원점수 ${sub.score}점, ${sub.rankRaw}위/${sub.totalStudents}명) -> 1등급 유지`);
        } else {
            lines.push(`${idx + 1}. ${sub.name}(${sub.units}단위): 현재 ${sub.currentGrade}등급 (원점수 ${sub.score}점, ${sub.rankRaw}위/${sub.totalStudents}명) -> 목표 ${sub.targetGrade}등급 (목표 ${sub.targetRank}위 이내, 앞에 ${sub.overtake}명 추월 필요)`);
        }
    });

    const clipboardText = lines.join('\n');
    navigator.clipboard.writeText(clipboardText).then(() => {
        const btn = document.getElementById('btnCopyAnalysisMid');
        if (!btn) return;
        const originalText = btn.innerHTML;
        btn.innerHTML = '<i class="fa-solid fa-check"></i> 복사 완료';
        btn.style.backgroundColor = '#0ca678';
        btn.style.color = 'white';
        btn.style.borderColor = '#0ca678';
        setTimeout(() => {
            btn.innerHTML = originalText;
            btn.style.backgroundColor = '';
            btn.style.color = '';
            btn.style.borderColor = '';
        }, 1500);
    });
}

// Copy Tab 1 All Students Analysis Results to Clipboard
function copyAllAnalysisMidToClipboard() {
    const studentIds = Object.keys(studentsDataMid).sort((a, b) => parseInt(a) - parseInt(b));
    if (studentIds.length === 0) {
        alert('업로드된 1회고사 학생 데이터가 없습니다.');
        return;
    }
    
    const allLines = [];
    studentIds.forEach(id => {
        const student = studentsDataMid[id];
        const subjectList = [];
        Object.keys(student.subjects).forEach(subName => {
            const subData = student.subjects[subName];
            const unitMatch = subName.match(/\((\d+)\)/);
            const units = unitMatch ? parseInt(unitMatch[1]) : 3;
            const { rank, ties } = parseRank(subData.rankRaw);
            if (rank === null || subData.totalStudents === null) return;

            const currentGrade = calculateGrade(rank, ties, subData.totalStudents);
            const targetGrade = Math.max(1, currentGrade - 1);
            const targetLimit = getTargetLimit(targetGrade);
            let targetRank = null;
            let overtake = 0;
            if (targetLimit !== null && currentGrade > 1) {
                targetRank = Math.round(subData.totalStudents * targetLimit);
                overtake = Math.max(0, rank - targetRank);
            }

            subjectList.push({
                name: subName.replace(/\(\d+\)/, "").trim(),
                units: units,
                score: subData.score,
                rankRaw: subData.rankRaw,
                totalStudents: subData.totalStudents,
                currentGrade: currentGrade,
                targetGrade: targetGrade,
                targetRank: targetRank,
                overtake: overtake
            });
        });

        subjectList.sort((a, b) => a.overtake - b.overtake);

        allLines.push(`[1회고사 성적 분석 결과 - ${student.name} (${id}번)]`);
        allLines.push(`평균 등급: ${calculateAverageGPA(student)}등급`);
        allLines.push(`과목별 2회고사 목표 및 추월 필요 인원 (우선순위 순):`);

        subjectList.forEach((sub, idx) => {
            if (sub.currentGrade === 1) {
                allLines.push(`${idx + 1}. ${sub.name}(${sub.units}단위): 현재 1등급 (원점수 ${sub.score}점, ${sub.rankRaw}위/${sub.totalStudents}명) -> 1등급 유지`);
            } else {
                allLines.push(`${idx + 1}. ${sub.name}(${sub.units}단위): 현재 ${sub.currentGrade}등급 (원점수 ${sub.score}점, ${sub.rankRaw}위/${sub.totalStudents}명) -> 목표 ${sub.targetGrade}등급 (목표 ${sub.targetRank}위 이내, 앞에 ${sub.overtake}명 추월 필요)`);
            }
        });
        allLines.push(`\n--------------------------------------------------\n`);
    });

    if (allLines.length > 0) {
        allLines.pop(); 
    }

    const clipboardText = allLines.join('\n');
    navigator.clipboard.writeText(clipboardText).then(() => {
        const btn = document.getElementById('btnCopyAllAnalysisMid');
        if (!btn) return;
        const originalText = btn.innerHTML;
        btn.innerHTML = '<i class="fa-solid fa-check"></i> 전체 복사 완료';
        btn.style.backgroundColor = '#0ca678';
        btn.style.color = 'white';
        btn.style.borderColor = '#0ca678';
        setTimeout(() => {
            btn.innerHTML = originalText;
            btn.style.backgroundColor = '';
            btn.style.color = '';
            btn.style.borderColor = '';
        }, 1500);
    });
}

function calculateAverageGPA(student) {
    let sumWeighted = 0;
    let sumUnits = 0;
    Object.keys(student.subjects).forEach(subName => {
        const subData = student.subjects[subName];
        const unitMatch = subName.match(/\((\d+)\)/);
        const units = unitMatch ? parseInt(unitMatch[1]) : 3;
        const { rank, ties } = parseRank(subData.rankRaw);
        if (rank === null || subData.totalStudents === null) return;
        const currentGrade = calculateGrade(rank, ties, subData.totalStudents);
        sumWeighted += currentGrade * units;
        sumUnits += units;
    });
    return sumUnits > 0 ? (sumWeighted / sumUnits).toFixed(2) : "-";
}

// Copy Tab 2 Semester Grade Table as TSV to Clipboard for Excel pasting
function copyTableToClipboard() {
    const student = studentsDataFull[selectedStudentId];
    if (!student) return;
    
    // Header row
    const headers = ['과목(단위수)'];
    const semesters = Object.keys(student.semesters).sort((a, b) => {
        const parse = (str) => {
            const m = str.match(/(\d+)학년\s+(\d+)학기/);
            return m ? parseInt(m[1]) * 10 + parseInt(m[2]) : 0;
        };
        return parse(a) - parse(b);
    });
    semesters.forEach(sem => headers.push(sem));
    
    const rows = [headers.join('\t')];
    
    // Subject rows
    const uniqueSubjects = [];
    semesters.forEach(sem => {
        student.semesters[sem]?.forEach(sub => {
            const exists = uniqueSubjects.find(s => s.name === sub.name);
            if (!exists) {
                uniqueSubjects.push({ name: sub.name, credits: sub.credits, category: sub.category });
            }
        });
    });
    
    // Sort subjects by category
    const catOrder = ['국어', '수학', '영어', '사회(역사/도덕포함)', '과학'];
    uniqueSubjects.sort((a, b) => {
        let idxA = catOrder.indexOf(a.category);
        let idxB = catOrder.indexOf(b.category);
        if (idxA === -1) idxA = 99;
        if (idxB === -1) idxB = 99;
        if (idxA !== idxB) return idxA - idxB;
        return a.name.localeCompare(b.name, 'ko');
    });
    
    uniqueSubjects.forEach(sub => {
        const rowCells = [`${sub.name}(${sub.credits})`];
        semesters.forEach(sem => {
            const matched = student.semesters[sem]?.find(s => s.name === sub.name);
            if (matched) {
                if (!isNaN(matched.grade)) {
                    rowCells.push(`${matched.grade}등급 (${matched.rawScore}점 / 성취도 ${matched.achievement} / ${matched.totalStudents ? matched.totalStudents + '명' : '-'})`);
                } else {
                    rowCells.push(`${matched.rawScore} (${matched.achievement ? '성취도 ' + matched.achievement : '-'} / ${matched.totalStudents ? matched.totalStudents + '명' : '-'})`);
                }
            } else {
                rowCells.push('-');
            }
        });
        rows.push(rowCells.join('\t'));
    });
    
    const clipboardText = rows.join('\n');
    navigator.clipboard.writeText(clipboardText).then(() => {
        const btn = document.getElementById('btnCopyTableFull');
        if (!btn) return;
        const originalText = btn.innerHTML;
        btn.innerHTML = '<i class="fa-solid fa-check"></i> 복사 완료';
        btn.style.backgroundColor = '#0ca678';
        btn.style.color = 'white';
        btn.style.borderColor = '#0ca678';
        setTimeout(() => {
            btn.innerHTML = originalText;
            btn.style.backgroundColor = '';
            btn.style.color = '';
            btn.style.borderColor = '';
        }, 1500);
    });
}
