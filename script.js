document.addEventListener('DOMContentLoaded', () => {
    // Check if SheetJS library is loaded and necessary functions exist
    // Now checking for sheet_to_json as the primary method
    if (typeof XLSX === 'undefined' || typeof XLSX.utils === 'undefined' || typeof XLSX.utils.sheet_to_json === 'undefined' || typeof XLSX.read === 'undefined') {
        console.error("SheetJS library (XLSX) or required functions (read, utils.sheet_to_json) not loaded correctly. Excel import/export might not work.");
        alert("警告：处理 Excel 的库未能正确加载或缺少必要功能，导入/导出功能可能无法使用。请检查网络连接或尝试刷新页面。");
        const importBtn = document.getElementById('import-options-btn');
        if(importBtn) importBtn.disabled = true;
        const exportNumBtn = document.getElementById('export-numbers-btn');
        // Keep export disabled as its functions might also be affected
    }

    // --- Random Number Generator (No changes) ---
    const minNumberInput = document.getElementById('min-number');
    const maxNumberInput = document.getElementById('max-number');
    const numberCountInput = document.getElementById('number-count');
    const generateNumberBtn = document.getElementById('generate-number-btn');
    const numberListUl = document.getElementById('random-number-list');
    const numberError = document.getElementById('number-error');
    const exportNumbersBtn = document.getElementById('export-numbers-btn');
    let currentGeneratedNumbers = [];
    generateNumberBtn.addEventListener('click', () => { /* ... same logic ... */
        const min = parseInt(minNumberInput.value); const max = parseInt(maxNumberInput.value); const count = parseInt(numberCountInput.value);
        numberError.textContent = ''; numberListUl.innerHTML = '<li>-</li>'; currentGeneratedNumbers = []; exportNumbersBtn.disabled = true;
        if (isNaN(min) || isNaN(max) || isNaN(count)) { numberError.textContent = '请输入有效的数字！'; return; }
        if (min > max) { numberError.textContent = '最小值不能大于最大值！'; return; }
        if (count <= 0) { numberError.textContent = '生成个数必须大于 0！'; return; }
        if (count > 5000) { numberError.textContent = '生成个数过多 (最多 5000)！'; return; }
        const randomNumbers = []; for (let i = 0; i < count; i++) { randomNumbers.push(Math.floor(Math.random() * (max - min + 1)) + min); }
        currentGeneratedNumbers = randomNumbers; numberListUl.innerHTML = '';
        currentGeneratedNumbers.forEach(num => { const li = document.createElement('li'); li.textContent = num; numberListUl.appendChild(li); });
        // Check for export function availability before enabling
        if (currentGeneratedNumbers.length > 0 && typeof XLSX !== 'undefined' && typeof XLSX.utils !== 'undefined' && typeof XLSX.utils.aoa_to_sheet !== 'undefined') { exportNumbersBtn.disabled = false; }
     });
    exportNumbersBtn.addEventListener('click', () => { /* ... same export logic using aoa_to_sheet ... */
        if (typeof XLSX === 'undefined' || typeof XLSX.utils === 'undefined' || typeof XLSX.utils.aoa_to_sheet === 'undefined' || typeof XLSX.utils.book_new === 'undefined' || typeof XLSX.utils.book_append_sheet === 'undefined' || typeof XLSX.writeFile === 'undefined') {
             alert("错误：Excel导出所需的功能未加载。"); return;
        }
        if (currentGeneratedNumbers.length === 0) { alert("没有数字可导出。"); return; } try { const data = [["Generated Random Numbers"], ...currentGeneratedNumbers.map(n => [n])]; const ws = XLSX.utils.aoa_to_sheet(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Random Numbers"); XLSX.writeFile(wb, "random_numbers.xlsx"); } catch (e) { console.error("导出数字错误:", e); alert("导出数字时出错。"); }
     });

    // --- Lucky Wheel ---
    const wheelSvg = document.getElementById('wheel-svg');
    const wheelContainer = document.querySelector('.wheel-container');
    const segmentsGroup = document.getElementById('wheel-segments-group');
    const textsGroup = document.getElementById('wheel-texts-group');
    const defs = wheelSvg.querySelector('defs');
    const spinWheelBtn = document.getElementById('spin-wheel-btn');
    const wheelResultSpan = document.getElementById('wheel-result');
    const wheelOptionsTextarea = document.getElementById('wheel-options');
    const updateWheelBtn = document.getElementById('update-wheel-btn');
    const wheelError = document.getElementById('wheel-error');
    const importOptionsBtn = document.getElementById('import-options-btn');
    const importFileInput = document.getElementById('import-file-input');

    const BLANK_SEGMENT_COUNT = 3;
    let currentSegments = [];
    const segmentColors = ['#FFD700', '#FF8C00', '#FF4500', '#DC143C', '#C71585', '#8A2BE2', '#4169E1', '#32CD32', '#00CED1', '#FF69B4'];
    let isSpinning = false;
    let svgRotation = 0;
    let isWheelBlank = false;

    const SVG_NS = "http://www.w3.org/2000/svg";
    const XLINK_NS = "http://www.w3.org/1999/xlink";
    const viewBoxSize = 200; const cx = viewBoxSize / 2; const cy = viewBoxSize / 2;
    const radius = (viewBoxSize / 2) * 0.95;

    function degToRad(degrees) { return degrees * (Math.PI / 180); }
    function getPoint(angleRad, r = radius) { return { x: cx + r * Math.cos(angleRad), y: cy + r * Math.sin(angleRad) }; }

    function renderWheelSVG() { /* ... SVG rendering logic remains exactly the same ... */
        segmentsGroup.innerHTML = ''; textsGroup.innerHTML = ''; defs.innerHTML = '';
        const segmentCount = currentSegments.length; if (segmentCount === 0) return;
        const angleDeg = 360 / segmentCount; const angleRad = degToRad(angleDeg);
        let wheelDisplaySize = 250; let baseFontSize = 10;
        if (segmentCount > 8) { wheelDisplaySize = 300; baseFontSize = 9; }
        if (segmentCount > 12) { wheelDisplaySize = 350; baseFontSize = 8; }
        wheelDisplaySize = Math.min(wheelDisplaySize, 400);
        wheelSvg.setAttribute('width', wheelDisplaySize); wheelSvg.setAttribute('height', wheelDisplaySize);
        wheelContainer.style.width = `${wheelDisplaySize}px`; wheelContainer.style.height = `${wheelDisplaySize}px`;
        for (let i = 0; i < segmentCount; i++) {
            const startAngleRad = angleRad * i - degToRad(90); const endAngleRad = startAngleRad + angleRad;
            const segmentText = currentSegments[i]; const segmentColor = segmentColors[i % segmentColors.length];
            const startPoint = getPoint(startAngleRad); const endPoint = getPoint(endAngleRad);
            const largeArcFlag = angleDeg > 180 ? 1 : 0;
            const pathData = `M ${cx},${cy} L ${startPoint.x},${startPoint.y} A ${radius},${radius} 0 ${largeArcFlag},1 ${endPoint.x},${endPoint.y} Z`;
            const segmentPath = document.createElementNS(SVG_NS, 'path');
            segmentPath.setAttribute('d', pathData); segmentPath.setAttribute('fill', segmentColor);
            segmentPath.classList.add('wheel-segment-path'); segmentsGroup.appendChild(segmentPath);
            if (segmentText !== "") {
                const textRadius = radius * 0.75;
                const textStartPoint = getPoint(startAngleRad, textRadius); const textEndPoint = getPoint(endAngleRad, textRadius);
                const textArcPath = document.createElementNS(SVG_NS, 'path'); const textPathId = `textArcPath_${i}`;
                textArcPath.setAttribute('id', textPathId);
                const textArcData = `M ${textStartPoint.x},${textStartPoint.y} A ${textRadius},${textRadius} 0 ${largeArcFlag},1 ${textEndPoint.x},${textEndPoint.y}`;
                textArcPath.setAttribute('d', textArcData); defs.appendChild(textArcPath);
                const textElement = document.createElementNS(SVG_NS, 'text');
                const textPathElement = document.createElementNS(SVG_NS, 'textPath');
                textPathElement.setAttributeNS(XLINK_NS, 'href', `#${textPathId}`);
                textPathElement.setAttribute('startOffset', '50%'); textPathElement.textContent = segmentText;
                textElement.appendChild(textPathElement); textElement.classList.add('wheel-segment-text');
                const fontSize = Math.max(6, baseFontSize - Math.floor(segmentCount / 10));
                textElement.setAttribute('font-size', `${fontSize}`); textsGroup.appendChild(textElement);
            }
        }
        svgRotation = 0; wheelSvg.style.transform = `rotate(${svgRotation}deg)`;
    }

    function updateWheel() { /* ... update logic remains the same ... */
        wheelError.textContent = ''; const optionsRaw = wheelOptionsTextarea.value; let options; isWheelBlank = false;
        if (optionsRaw.trim() === '') { options = Array(BLANK_SEGMENT_COUNT).fill(""); isWheelBlank = true; }
        else { options = optionsRaw.split('\n').map(o => o.trim()); }
        const minReq = isWheelBlank ? BLANK_SEGMENT_COUNT : 2;
        if (options.length < minReq) { wheelError.textContent = `转盘至少需要 ${minReq} 个扇区。`; spinWheelBtn.disabled = true; segmentsGroup.innerHTML = ''; textsGroup.innerHTML = ''; defs.innerHTML = ''; currentSegments = []; return; }
        currentSegments = options; renderWheelSVG(); spinWheelBtn.disabled = false; importOptionsBtn.disabled = (typeof XLSX === 'undefined'); wheelResultSpan.textContent = '-'; wheelResultSpan.style.backgroundColor = ''; wheelResultSpan.style.color = '';
    }

    updateWheelBtn.addEventListener('click', updateWheel);

    spinWheelBtn.addEventListener('click', () => { /* ... spin logic remains the same ... */
        if (isSpinning || currentSegments.length < 2) return; isSpinning = true; spinWheelBtn.disabled = true; updateWheelBtn.disabled = true; importOptionsBtn.disabled = true; wheelResultSpan.textContent = '...'; wheelResultSpan.style.backgroundColor = ''; wheelResultSpan.style.color = ''; const segmentCount = currentSegments.length; const degreesPerSegment = 360 / segmentCount; const randomSegmentIndex = Math.floor(Math.random() * segmentCount); const targetMiddleAngle = randomSegmentIndex * degreesPerSegment + degreesPerSegment / 2; let stoppingAngle = 360 - targetMiddleAngle; const extraRotations = Math.floor(Math.random() * 6) + 5; const currentVisualAngle = parseFloat(wheelSvg.style.transform.replace(/[^-\d.]/g, '')) || 0; svgRotation = currentVisualAngle + (360*extraRotations) + (stoppingAngle - (currentVisualAngle % 360) + 360) % 360; wheelSvg.style.transform = `rotate(${svgRotation}deg)`; setTimeout(() => { const finalVisualAngle = svgRotation % 360; const normalizedAngle = (360 - finalVisualAngle) % 360; const winningIndex = Math.floor(normalizedAngle / degreesPerSegment) % segmentCount; if (isWheelBlank) { const winningColor = segmentColors[winningIndex % segmentColors.length]; wheelResultSpan.textContent = `${winningColor}`; wheelResultSpan.style.backgroundColor = winningColor; const darkColors = ['#DC143C', '#C71585', '#8A2BE2', '#4169E1']; wheelResultSpan.style.color = darkColors.includes(winningColor) ? '#fff' : '#333'; } else { wheelResultSpan.textContent = currentSegments[winningIndex]; wheelResultSpan.style.backgroundColor = ''; wheelResultSpan.style.color = ''; } isSpinning = false; if (currentSegments.length >= 2) { spinWheelBtn.disabled = false; updateWheelBtn.disabled = false; importOptionsBtn.disabled = (typeof XLSX === 'undefined'); } }, 4100);
     });

    importOptionsBtn.addEventListener('click', () => {
        if (typeof XLSX === 'undefined' || typeof FileReader === 'undefined') { alert("错误：文件处理所需的功能未加载或浏览器不支持。"); return; }
        importFileInput.click();
    });

    importFileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) { return; }
        const reader = new FileReader();
        reader.onload = (e) => {
            // --- 诊断代码 (更新) ---
            console.log('尝试处理文件... 读取完成。');
            console.log('XLSX 对象类型:', typeof XLSX);
            if (typeof XLSX !== 'undefined') {
                console.log('XLSX.utils 对象类型:', typeof XLSX.utils);
                if (typeof XLSX.utils !== 'undefined') {
                    console.log('XLSX.utils.sheet_to_aoa 函数类型:', typeof XLSX.utils.sheet_to_aoa); // << 还是检查一下 aoa
                    console.log('XLSX.utils.sheet_to_json 函数类型:', typeof XLSX.utils.sheet_to_json); // << 检查 json
                } else { console.log('错误：XLSX.utils 未定义！'); }
            } else { console.log('错误：XLSX 对象未定义！库未加载？'); }
            // --- 诊断代码结束 ---

            // Defensive check for required functions (now check sheet_to_json)
             if (typeof XLSX === 'undefined' || typeof XLSX.read === 'undefined' || typeof XLSX.utils === 'undefined' || typeof XLSX.utils.sheet_to_json === 'undefined') {
                 alert("错误：必需的 Excel 处理函数(sheet_to_json)缺失。库可能加载不完整或版本错误。");
                 event.target.value = null; return;
             }

            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                if (!firstSheetName) { throw new Error("文件中没有找到工作表。"); }
                const worksheet = workbook.Sheets[firstSheetName];
                if (!worksheet) { throw new Error(`找不到名为 "${firstSheetName}" 的工作表。`); }

                // *** MODIFIED: Use sheet_to_json instead of sheet_to_aoa ***
                console.log("尝试使用 XLSX.utils.sheet_to_json...");
                const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                console.log("sheet_to_json 执行成功!");
                // *** End of Modification ***

                const options = sheetData
                    .map(row => (row[0] !== null && row[0] !== undefined) ? String(row[0]).trim() : "")
                    .filter((value, index, arr) => value !== "" || arr.slice(index + 1).some(v => v !== ""));

                if (options.length === 0) {
                    alert("在文件的第一列中未找到有效选项。");
                    return;
                }
                wheelOptionsTextarea.value = options.join('\n');
                updateWheel();

            } catch (error) {
                console.error("处理文件时出错:", error);
                // Check if the error is the specific "not a function" error again, even for sheet_to_json
                if (error instanceof TypeError && error.message.includes("is not a function")) {
                     alert(`处理文件时出错: ${error.message}\n库文件可能仍有问题或浏览器环境冲突。`);
                } else {
                     alert(`处理文件时出错: ${error.message}\n请确保文件格式正确 (Excel/CSV，选项在第一列)。`);
                }
            } finally {
                 event.target.value = null;
            }
        };
        reader.onerror = (e) => {
            alert('读取文件时出错。'); console.error("FileReader error:", e);
            event.target.value = null;
        };
        reader.readAsArrayBuffer(file);
    });

    // Initial Setup
    updateWheel();
});