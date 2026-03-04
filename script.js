// 全局变量
let isRolling = false; // 是否正在抽奖
let rollingInterval; // 抽奖滚动定时器
let currentAward = null; // 当前选中的奖项
let awardParticipants = []; // 当前奖项的可选参与人员
let lotteryData = {
    activityName: '',
    activityBackground: null,
    awards: [],
    participants: [],
    winners: []
};

// DOM元素
const settingsPanel = document.getElementById('settingsPanel');
const toggleBtn = document.getElementById('toggleBtn');
const toggleIcon = document.getElementById('toggleIcon');
const panelContent = document.getElementById('panelContent');
const activityNameInput = document.getElementById('activityName');
const activityBackgroundInput = document.getElementById('activityBackground');
const awardsTable = document.getElementById('awardsTable');
const participantsTable = document.getElementById('participantsTable');
const resetBtn = document.getElementById('resetBtn');
const saveBtn = document.getElementById('saveBtn');
const defaultState = document.getElementById('defaultState');
const normalState = document.getElementById('normalState');
const activityTitle = document.getElementById('activityTitle');
const awardSelect = document.getElementById('awardSelect');
const selectedAwardImage = document.getElementById('selectedAwardImage');
const selectedAwardName = document.getElementById('selectedAwardName');
const awardQuota = document.getElementById('awardQuota');
const remainingQuota = document.getElementById('remainingQuota');
const rollingNames = document.getElementById('rollingNames');
const startBtn = document.getElementById('startBtn');
const stopBtn = document.getElementById('stopBtn');
const resetAwardBtn = document.getElementById('resetAwardBtn');
const recordsContainer = document.getElementById('recordsContainer');

// 弹窗元素
const resetModal = document.getElementById('resetModal');
const cancelResetBtn = document.getElementById('cancelResetBtn');
const confirmResetBtn = document.getElementById('confirmResetBtn');
const validationModal = document.getElementById('validationModal');
const validationMessage = document.getElementById('validationMessage');
const closeValidationBtn = document.getElementById('closeValidationBtn');
const successModal = document.getElementById('successModal');
const closeSuccessBtn = document.getElementById('closeSuccessBtn');
const importBtn = document.getElementById('importBtn');
const excelFile = document.getElementById('excelFile');

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    // 初始化表格事件监听
    initTableEvents();
    
    // 初始化按钮事件监听
    initButtonEvents();
    
    // 初始化弹窗事件监听
    initModalEvents();
    
    // 添加键盘事件监听器，支持空格键控制抽奖
    document.addEventListener('keydown', function(event) {
        // 检查是否按下了空格键
        if (event.code === 'Space') {
            // 阻止默认的空格键行为（如页面滚动）
            event.preventDefault();
            
            // 根据当前状态执行相应操作
            if (isRolling) {
                stopLottery();
            } else {
                startLottery();
            }
        }
    });
});

// 初始化表格事件监听
function initTableEvents() {
    // 奖项表格事件委托
    awardsTable.addEventListener('click', function(e) {
        if (e.target.classList.contains('add-row-btn')) {
            addTableRow(e.target.closest('tr'), awardsTable);
        } else if (e.target.classList.contains('delete-row-btn')) {
            deleteTableRow(e.target.closest('tr'), awardsTable);
        }
    });
    
    // 名单表格事件委托
    participantsTable.addEventListener('click', function(e) {
        if (e.target.classList.contains('add-row-btn')) {
            addTableRow(e.target.closest('tr'), participantsTable);
        } else if (e.target.classList.contains('delete-row-btn')) {
            deleteTableRow(e.target.closest('tr'), participantsTable);
        }
    });
    
    // 奖项图片预览
    awardsTable.addEventListener('change', function(e) {
        if (e.target.classList.contains('award-image')) {
            previewImage(e.target);
        }
    });
    
    // 活动背景图预览
    activityBackgroundInput.addEventListener('change', function() {
        previewBackground();
    });
}

// 初始化按钮事件监听
function initButtonEvents() {
    // 面板收齐/展开按钮
    toggleBtn.addEventListener('click', function() {
        settingsPanel.classList.toggle('collapsed');
        toggleIcon.textContent = settingsPanel.classList.contains('collapsed') ? '→' : '←';
    });
    
    // 重置设置按钮
    resetBtn.addEventListener('click', function() {
        resetModal.style.display = 'flex';
    });
    
    // 保存设置按钮
    saveBtn.addEventListener('click', function() {
        saveSettings();
    });
    
    // 导入按钮点击事件
    importBtn.addEventListener('click', function() {
        excelFile.click();
    });
    
    // 文件选择事件
    excelFile.addEventListener('change', handleFileImport);
    
    // 奖项选择器变更
    awardSelect.addEventListener('change', function() {
        updateAwardInfo();
    });
    
    // 开始抽奖按钮
    startBtn.addEventListener('click', function() {
        startLottery();
    });
    
    // 停止抽奖按钮
    stopBtn.addEventListener('click', function() {
        stopLottery();
    });
    
    // 重置抽奖按钮
    resetAwardBtn.addEventListener('click', function() {
        resetAwardLottery();
    });
    
    // 导出Excel按钮
    document.getElementById('exportBtn').addEventListener('click', function() {
        exportToExcel();
    });
    
    // 全屏按钮
    const fullscreenBtn = document.getElementById('fullscreenBtn');
    if (fullscreenBtn) {
        fullscreenBtn.addEventListener('click', toggleFullscreen);
    }
}

// 切换全屏
function toggleFullscreen() {
    const lotteryPanel = document.querySelector('.lottery-panel');
    
    if (!document.fullscreenElement) {
        // 进入全屏
        if (lotteryPanel.requestFullscreen) {
            lotteryPanel.requestFullscreen();
        } else if (lotteryPanel.webkitRequestFullscreen) {
            lotteryPanel.webkitRequestFullscreen();
        } else if (lotteryPanel.msRequestFullscreen) {
            lotteryPanel.msRequestFullscreen();
        }
    } else {
        // 退出全屏
        if (document.exitFullscreen) {
            document.exitFullscreen();
        } else if (document.webkitExitFullscreen) {
            document.webkitExitFullscreen();
        } else if (document.msExitFullscreen) {
            document.msExitFullscreen();
        }
    }
}

// 初始化弹窗事件监听
function initModalEvents() {
    // 重置确认弹窗
    cancelResetBtn.addEventListener('click', function() {
        resetModal.style.display = 'none';
    });
    
    confirmResetBtn.addEventListener('click', function() {
        resetSettings();
        resetModal.style.display = 'none';
    });
    
    // 表单验证提示弹窗
    closeValidationBtn.addEventListener('click', function() {
        validationModal.style.display = 'none';
    });
    
    // 保存成功提示弹窗
    closeSuccessBtn.addEventListener('click', function() {
        successModal.style.display = 'none';
    });
    
    // 点击弹窗外部关闭弹窗
    window.addEventListener('click', function(e) {
        if (e.target.classList.contains('modal')) {
            e.target.style.display = 'none';
        }
    });
}

// 添加表格行
function addTableRow(currentRow, table) {
    const newRow = document.createElement('tr');
    let rowHTML = '';
    
    if (table.id === 'awardsTable') {
        rowHTML = `
            <td>
                <input type="text" class="award-name" placeholder="请输入奖项名称" maxlength="30">
            </td>
            <td>
                <input type="file" class="award-image" accept="image/*">
            </td>
            <td>
                <input type="number" class="award-quota" min="1" max="100" value="1" step="1" onchange="ensureMinValue(this)">
            </td>
            <td>
                <button class="add-row-btn">+</button>
                <button class="delete-row-btn">-</button>
            </td>
        `;
    } else if (table.id === 'participantsTable') {
        rowHTML = `
            <td>
                <input type="text" class="participant-name" placeholder="请输入人员姓名" maxlength="30">
            </td>
            <td>
                <input type="number" class="participant-count" min="1" max="100" value="1" step="1" onchange="ensureMinValue(this)">
            </td>
            <td>
                <button class="add-row-btn">+</button>
                <button class="delete-row-btn">-</button>
            </td>
        `;
    }
    
    newRow.innerHTML = rowHTML;
    currentRow.parentNode.insertBefore(newRow, currentRow);
    
    // 更新删除按钮状态
    updateDeleteButtonState(table);
    
    // 为新添加的行绑定事件
    if (table.id === 'awardsTable') {
        const imageInput = newRow.querySelector('.award-image');
        imageInput.addEventListener('change', function() {
            previewImage(imageInput);
        });
    }
}

// 删除表格行
function deleteTableRow(row, table) {
    if (table.querySelectorAll('tbody tr').length > 1) {
        row.remove();
        updateDeleteButtonState(table);
    }
}

// 更新删除按钮状态
function updateDeleteButtonState(table) {
    const rows = table.querySelectorAll('tbody tr');
    const deleteButtons = table.querySelectorAll('.delete-row-btn');
    
    deleteButtons.forEach(btn => {
        btn.disabled = rows.length <= 1;
    });
}

// 预览图片
function previewImage(input) {
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            // 隐藏文件输入框
            input.style.display = 'none';
            
            // 创建或获取预览元素
            let previewContainer = input.parentNode.querySelector('.image-preview-container');
            if (!previewContainer) {
                previewContainer = document.createElement('div');
                previewContainer.className = 'image-preview-container';
                
                // 创建预览图片元素
                const previewImg = document.createElement('img');
                previewImg.className = 'preview-image';
                previewImg.style.cssText = `
                    max-width: 100%;
                    max-height: 100%;
                    display: block;
                `;
                
                // 创建覆盖层
                const overlay = document.createElement('div');
                overlay.className = 'image-overlay';
                overlay.style.cssText = `
                    position: absolute;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    background-color: rgba(0, 0, 0, 0.6);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    opacity: 0;
                    transition: opacity 0.3s;
                `;
                
                // 创建替换按钮
                const replaceBtn = document.createElement('button');
                replaceBtn.type = 'button';
                replaceBtn.textContent = '替换图片';
                replaceBtn.style.cssText = `
                    background-color: #4a90e2;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                `;
                
                // 点击替换按钮重新触发文件输入
                replaceBtn.onclick = function(e) {
                    e.stopPropagation();
                    input.click(); // 触发原始文件输入框
                };
                
                overlay.appendChild(replaceBtn);
                
                // 将图片和覆盖层放入容器
                previewContainer.style.cssText = `
                    position: relative;
                    width: 100px;
                    height: 100px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    overflow: hidden;
                `;
                
                previewContainer.appendChild(previewImg);
                previewContainer.appendChild(overlay);
                
                // 鼠标悬停显示覆盖层
                previewContainer.onmouseover = function() {
                    overlay.style.opacity = '1';
                };
                
                previewContainer.onmouseout = function() {
                    overlay.style.opacity = '0';
                };
                
                input.parentNode.style.position = 'relative';
                input.parentNode.appendChild(previewContainer);
            }
            
            // 更新预览图片源
            previewContainer.querySelector('.preview-image').src = e.target.result;
            
            // 保存Data URL到input元素的自定义属性中
            input.dataset.imageUrl = e.target.result;
        };
        
        reader.readAsDataURL(input.files[0]);
    }
}

// 预览活动背景图
function previewBackground() {
    if (activityBackgroundInput.files && activityBackgroundInput.files[0]) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            const previewContainer = document.getElementById('backgroundPreviewContainer');
            
            // 创建预览图片元素
            let previewImg = document.getElementById('backgroundPreview');
            if (!previewImg) {
                previewImg = document.createElement('img');
                previewImg.id = 'backgroundPreview';
                previewImg.alt = '背景预览';
                previewContainer.insertBefore(previewImg, document.getElementById('backgroundOverlay'));
            }
            
            previewImg.src = e.target.result;
            previewContainer.style.display = 'block';
            previewContainer.classList.add('has-image'); // 添加类以隐藏缺省提示
            
            // 保存Data URL到input元素的自定义属性中
            activityBackgroundInput.dataset.imageUrl = e.target.result;
        };
        
        reader.readAsDataURL(activityBackgroundInput.files[0]);
    }
}

// 替换背景图
function replaceBackground() {
    activityBackgroundInput.click();
}

// 重置设置
function resetSettings() {
    // 重置活动名称
    activityNameInput.value = '';
    
    // 重置活动背景图
    activityBackgroundInput.value = '';
    const backgroundPreviewContainer = document.getElementById('backgroundPreviewContainer');
    backgroundPreviewContainer.style.display = 'block';
    backgroundPreviewContainer.classList.remove('has-image'); // 移除类以显示缺省提示
    
    // 移除预览图片元素
    const backgroundPreview = document.getElementById('backgroundPreview');
    if (backgroundPreview) {
        backgroundPreview.remove();
    }
    
    // 重置奖项表格
    const awardsRows = awardsTable.querySelectorAll('tbody tr');
    awardsRows.forEach((row, index) => {
        if (index > 0) {
            row.remove();
        } else {
            row.querySelector('.award-name').value = '';
            row.querySelector('.award-image').value = '';
            row.querySelector('.award-quota').value = '1';
            
            // 移除预览容器和显示的文件输入框
            const previewContainer = row.querySelector('.image-preview-container');
            if (previewContainer) {
                previewContainer.remove();
            }
            
            // 显示隐藏的文件输入框
            const imageInput = row.querySelector('.award-image');
            imageInput.style.display = 'block'; // 重新显示文件输入框
        }
    });
    
    // 重置名单表格
    const participantsRows = participantsTable.querySelectorAll('tbody tr');
    participantsRows.forEach((row, index) => {
        if (index > 0) {
            row.remove();
        } else {
            row.querySelector('.participant-name').value = '';
            row.querySelector('.participant-count').value = '1';
        }
    });
    
    // 更新删除按钮状态
    updateDeleteButtonState(awardsTable);
    updateDeleteButtonState(participantsTable);
    
    // 重置抽奖数据
    lotteryData = {
        activityName: '',
        activityBackground: null,
        awards: [],
        participants: [],
        winners: []
    };
    
    // 重置右侧抽奖大屏
    defaultState.style.display = 'flex';
    normalState.style.display = 'none';
    
    // 重置开始/停止按钮状态
    startBtn.disabled = true;
    stopBtn.disabled = true;
}

// 保存设置
async function saveSettings() {
    // 验证活动名称
    const activityName = activityNameInput.value.trim();
    if (!activityName) {
        showValidationError('请输入活动名称');
        return;
    }
    
    // 验证奖项设置
    const awardsRows = awardsTable.querySelectorAll('tbody tr');
    const awards = [];
    
    for (let row of awardsRows) {
        const name = row.querySelector('.award-name').value.trim();
        const imageInput = row.querySelector('.award-image');
        const quota = parseInt(row.querySelector('.award-quota').value);
        
        if (!name) {
            showValidationError('请填写所有奖项名称');
            return;
        }
        
        if (!imageInput.files || !imageInput.files[0]) {
            showValidationError('请为所有奖项上传图片');
            return;
        }
        
        if (isNaN(quota) || quota < 1 || quota > 100) {
            showValidationError('奖项名额必须在1-100之间');
            return;
        }
        
        // 使用之前保存的Data URL，或者从文件读取
        let imageDataUrl = imageInput.dataset.imageUrl;
        if (!imageDataUrl) {
            // 如果没有预览URL，则从文件创建
            imageDataUrl = URL.createObjectURL(imageInput.files[0]);
        }
        
        const awardObj = {
            name,
            image: imageInput.files[0],
            quota,
            remainingQuota: quota,
            imageDataUrl: imageDataUrl
        };
        
        awards.push(awardObj);
    }
    
    // 验证名单设置
    const participantsRows = participantsTable.querySelectorAll('tbody tr');
    const participants = [];
    const participantNames = [];
    
    for (let row of participantsRows) {
        const name = row.querySelector('.participant-name').value.trim();
        const count = parseInt(row.querySelector('.participant-count').value);
        
        if (!name) {
            showValidationError('请填写所有人员姓名');
            return;
        }
        
        if (participantNames.includes(name)) {
            showValidationError('人员姓名不能重复');
            return;
        }
        
        if (isNaN(count) || count < 1 || count > 100) {
            showValidationError('抽奖票数必须在1-100之间');
            return;
        }
        
        participantNames.push(name);
        participants.push({
            name,
            count,
            availableCount: count // 可中奖次数
        });
    }
    
    // 获取活动背景图数据
    let activityBackgroundUrl = null;
    if (activityBackgroundInput && activityBackgroundInput.dataset.imageUrl) {
        activityBackgroundUrl = activityBackgroundInput.dataset.imageUrl;
    }
    
    // 更新抽奖数据
    lotteryData.activityName = activityName;
    lotteryData.activityBackground = activityBackgroundUrl;
    lotteryData.awards = awards;
    lotteryData.participants = participants;
    
    // 更新右侧抽奖大屏
    updateLotteryPanel();
    
    // 显示保存成功弹窗
    successModal.style.display = 'flex';
}

// 更新抽奖大屏
function updateLotteryPanel() {
    // 显示正常状态，隐藏缺省状态
    defaultState.style.display = 'none';
    normalState.style.display = 'block';
    
    // 更新活动名称
    activityTitle.textContent = lotteryData.activityName;
    
    // 设置背景图
    const lotteryPanel = document.querySelector('.lottery-panel');
    if (lotteryData.activityBackground) {
        lotteryPanel.style.backgroundImage = `url('${lotteryData.activityBackground}')`;
        lotteryPanel.style.backgroundSize = 'cover';
        lotteryPanel.style.backgroundPosition = 'center';
        lotteryPanel.style.backgroundColor = '#f9f9f9'; // 备用颜色
    } else {
        lotteryPanel.style.backgroundImage = 'none';
        lotteryPanel.style.backgroundColor = '#f9f9f9';
    }
    
    // 更新奖项选择器
    awardSelect.innerHTML = '';
    lotteryData.awards.forEach((award, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = award.name;
        awardSelect.appendChild(option);
    });
    
    // 更新奖项信息
    updateAwardInfo();
    
    // 清空中奖记录
    recordsContainer.innerHTML = '';
    lotteryData.winners.forEach(winner => {
        addWinnerRecord(winner);
    });
}

// 更新奖项信息
function updateAwardInfo() {
    const awardIndex = parseInt(awardSelect.value);
    currentAward = lotteryData.awards[awardIndex];
    
    if (currentAward) {
        // 直接使用已存储的DataURL更新奖项图片
        selectedAwardImage.src = currentAward.imageDataUrl;
        
        // 更新奖项名称和名额
        selectedAwardName.textContent = currentAward.name;
        awardQuota.textContent = currentAward.quota;
        remainingQuota.textContent = currentAward.remainingQuota;
        
        // 更新开始抽奖按钮状态
        startBtn.disabled = currentAward.remainingQuota <= 0;
        stopBtn.disabled = true;
        resetAwardBtn.disabled = false;
        
        // 重置抽奖状态
        rollingNames.textContent = '请点击【开始抽奖】';
    }
    
    // 重置抽奖状态
    isRolling = false;
    clearInterval(rollingInterval);
    stopBtn.disabled = true;
    resetAwardBtn.disabled = false;
}

// 开始抽奖
function startLottery() {
    if (!currentAward || currentAward.remainingQuota <= 0) {
        return;
    }
    
    // 准备可选参与人员
    prepareAwardParticipants();
    
    if (awardParticipants.length === 0) {
        rollingNames.textContent = '没有可选人员';
        return;
    }
    
    // 开始滚动
    isRolling = true;
    startBtn.disabled = true;
    stopBtn.disabled = false;
    
    // 播放抽奖音乐
    try {
        const lotteryMusic = document.getElementById('lotteryMusic');
        if (lotteryMusic) {
            // 确保音频已加载
            lotteryMusic.load();
            lotteryMusic.volume = 0.8;
            
            // 尝试播放并处理可能的错误
            const playPromise = lotteryMusic.play();
            if (playPromise !== undefined) {
                playPromise.then(() => {
                    console.log('音乐播放成功');
                }).catch(error => {
                    console.warn('音乐自动播放被阻止，需要用户交互:', error);
                    // 添加点击事件以便用户可以手动播放
                    document.addEventListener('click', function enableAudio() {
                        lotteryMusic.play().then(() => {
                            console.log('音乐通过用户交互开始播放');
                        }).catch(e => {
                            console.log('音乐播放失败:', e);
                        });
                        document.removeEventListener('click', enableAudio);
                    }, { once: true });
                });
            }
        }
    } catch (error) {
        console.log('处理音频时出错:', error);
    }
    
    rollingInterval = setInterval(() => {
        const randomIndex = Math.floor(Math.random() * awardParticipants.length);
        rollingNames.textContent = awardParticipants[randomIndex].name;
    }, 50);
}

// 准备可选参与人员
function prepareAwardParticipants() {
    awardParticipants = [];
    
    // 筛选可参与当前奖项的人员
    lotteryData.participants.forEach(participant => {
        // 检查该人员是否已经中过当前奖项
        const hasWonCurrentAward = lotteryData.winners.some(winner => 
            winner.name === participant.name && winner.awardName === currentAward.name
        );
        
        // 如果该人员还有可用次数且未中过当前奖项，则添加到可选列表
        if (participant.availableCount > 0 && !hasWonCurrentAward) {
            // 根据可用次数添加到可选列表
            for (let i = 0; i < participant.availableCount; i++) {
                awardParticipants.push(participant);
            }
        }
    });
}

// 停止抽奖
function stopLottery() {
    if (!isRolling) {
        return;
    }
    
    // 停止滚动
    isRolling = false;
    clearInterval(rollingInterval);
    stopBtn.disabled = true;
    
    // 暂停抽奖音乐
    try {
        const lotteryMusic = document.getElementById('lotteryMusic');
        if (lotteryMusic) {
            lotteryMusic.pause();
        }
    } catch (error) {
        console.log('处理音频时出错:', error);
        // 音频错误不影响抽奖流程
    }
    
    // 确定中奖人员
    const winner = determineWinner();
    
    if (winner) {
        // 显示中奖动画效果
        rollingNames.textContent = winner.name;
        rollingNames.classList.add('win-animation');
        
        // 记录中奖信息
        const winnerInfo = {
            name: winner.name,
            awardName: currentAward.name,
            time: new Date()
        };
        
        lotteryData.winners.push(winnerInfo);
        
        // 更新参与人员可用次数
        const participant = lotteryData.participants.find(p => p.name === winner.name);
        if (participant) {
            participant.availableCount--;
        }
        
        // 更新奖项剩余名额
        currentAward.remainingQuota--;
        remainingQuota.textContent = currentAward.remainingQuota;
        
        // 添加中奖记录
        addWinnerRecord(winnerInfo);
        
        // 更新开始抽奖按钮状态
        startBtn.disabled = currentAward.remainingQuota <= 0;
        
        // 显示中奖动效和彩带飘落
        showWinEffect(winnerInfo);
        
        // 1秒后移除动画类，以便下次抽奖时可以再次触发
        setTimeout(() => {
            rollingNames.classList.remove('win-animation');
        }, 1000);
    }
}

// 显示中奖动效和彩带飘落
function showWinEffect(winnerInfo) {
    const winEffectContainer = document.getElementById('winEffectContainer');
    const winName = document.getElementById('winName');
    const winAward = document.getElementById('winAward');
    const confettiContainer = document.getElementById('confettiContainer');
    
    // 设置中奖信息
    winName.textContent = `中奖人员：${winnerInfo.name}`;
    winAward.textContent = `中奖奖项：${winnerInfo.awardName}`;
    
    // 显示中奖动效容器
    winEffectContainer.classList.add('show');
    
    // 生成彩带
    createConfetti(confettiContainer);
    
    // 5秒后隐藏中奖动效容器
    setTimeout(() => {
        winEffectContainer.classList.remove('show');
        // 清空彩带容器
        confettiContainer.innerHTML = '';
    }, 5000);
}

// 生成彩带
function createConfetti(container) {
    // 彩带动画持续时间
    const duration = 3000;
    // 生成彩带数量
    const count = 100;
    // 彩带颜色
    const colors = ['#f00', '#0f0', '#00f', '#ff0', '#f0f', '#0ff'];
    
    // 生成彩带
    for (let i = 0; i < count; i++) {
        const confetti = document.createElement('div');
        confetti.className = 'confetti';
        
        // 随机设置彩带颜色
        confetti.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
        
        // 随机设置彩带大小
        const size = Math.random() * 10 + 5;
        confetti.style.width = `${size}px`;
        confetti.style.height = `${size}px`;
        
        // 随机设置彩带位置
        confetti.style.left = `${Math.random() * 100}vw`;
        confetti.style.top = '-20px';
        
        // 随机设置彩带动画持续时间
        const animationDuration = Math.random() * 2 + 2;
        confetti.style.animationDuration = `${animationDuration}s`;
        
        // 随机设置彩带动画延迟
        confetti.style.animationDelay = `${Math.random() * 1}s`;
        
        // 添加彩带到容器
        container.appendChild(confetti);
    }
}

// 确定中奖人员
function determineWinner() {
    if (awardParticipants.length === 0) {
        return null;
    }
    
    const randomIndex = Math.floor(Math.random() * awardParticipants.length);
    return awardParticipants[randomIndex];
}

// 添加中奖记录
function addWinnerRecord(winner) {
    const recordItem = document.createElement('div');
    recordItem.className = 'record-item';
    
    const timeStr = formatDateTime(winner.time);
    
    recordItem.innerHTML = `
        <span class="record-name">${winner.name}</span>
        <span class="record-award">${winner.awardName}</span>
        <span class="record-time">${timeStr}</span>
    `;
    
    recordsContainer.insertBefore(recordItem, recordsContainer.firstChild);
}

// 格式化日期时间
function formatDateTime(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// 显示表单验证错误
function showValidationError(message) {
    validationMessage.textContent = message;
    validationModal.style.display = 'flex';
}

// 确保最小值为1的函数
function ensureMinValue(input) {
    if (input.value === '' || parseInt(input.value) < 1) {
        input.value = 1;
    }
}

// 处理文件导入
function handleFileImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const fileName = file.name;
    const fileExtension = fileName.split('.').pop().toLowerCase();
    
    if (fileExtension === 'csv') {
        // 处理CSV文件
        handleCSVFile(file);
    } else if (['xlsx', 'xls'].includes(fileExtension)) {
        // 处理Excel文件
        handleExcelFile(file);
    } else {
        showValidationError('仅支持CSV、XLS和XLSX格式的文件');
    }
    
    // 重置文件输入，允许重复选择同一个文件
    e.target.value = '';
}

// 处理CSV文件
function handleCSVFile(file) {
    // 由于Windows下Excel导出的CSV默认使用GBK编码，而FileReader不直接支持GBK
    // 我们使用readAsArrayBuffer来读取文件，然后尝试多种编码解析
    const reader = new FileReader();
    reader.onload = function(e) {
        const arrayBuffer = e.target.result;
        const uint8Array = new Uint8Array(arrayBuffer);
        
        // 尝试多种编码
        const encodings = ['UTF-8', 'GBK', 'GB18030', 'ISO-8859-1'];
        let content = '';
        let encodingSuccess = false;
        
        for (const encoding of encodings) {
            try {
                content = new TextDecoder(encoding).decode(uint8Array);
                // 检查解码是否成功（简单判断：如果不包含大量连续的特殊字符，可能解码成功）
                const hasGibberish = /[\?�]{3,}/.test(content);
                if (!hasGibberish) {
                    encodingSuccess = true;
                    break;
                }
            } catch (e) {
                // 编码不支持，继续尝试下一个
                continue;
            }
        }
        
        // 如果所有编码都失败，使用默认编码
        if (!encodingSuccess) {
            content = new TextDecoder().decode(uint8Array);
        }
        
        // 处理UTF-8 BOM
        if (content.charCodeAt(0) === 0xFEFF) {
            content = content.slice(1);
        }
        
        const lines = content.split('\n');
        
        // 跳过表头
        const data = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            
            let values = [];
            
            // 首先尝试按制表符分割（常见于Excel导出的CSV）
            if (line.includes('\t')) {
                values = line.split('\t').map(val => val.trim());
            } else {
                // 按逗号分割，处理包含逗号的字段
                values = [];
                let currentValue = '';
                let inQuotes = false;
                
                for (let j = 0; j < line.length; j++) {
                    const char = line[j];
                    if (char === '"') {
                        inQuotes = !inQuotes;
                    } else if (char === ',' && !inQuotes) {
                        values.push(currentValue);
                        currentValue = '';
                    } else {
                        currentValue += char;
                    }
                }
                values.push(currentValue);
                values = values.map(val => val.trim());
            }
            
            if (values.length >= 2) {
                const name = values[0];
                const countStr = values[1];
                const count = parseInt(countStr);
                
                if (name && !isNaN(count) && count >= 1 && count <= 100) {
                    data.push({ name, count });
                }
            }
        }
        
        if (data.length > 0) {
            importDataToTable(data);
        } else {
            showValidationError('未找到有效的数据，请检查文件格式');
        }
    };
    reader.onerror = function() {
        showValidationError('文件读取失败');
    };
    reader.readAsArrayBuffer(file);
}

// 处理Excel文件
function handleExcelFile(file) {
    // 检查是否已加载SheetJS库
    if (typeof XLSX === 'undefined') {
        // 如果没有SheetJS库，尝试动态加载
        loadSheetJS(() => {
            parseExcelFile(file);
        });
    } else {
        parseExcelFile(file);
    }
}

// 加载SheetJS库
function loadSheetJS(callback) {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    script.onload = callback;
    script.onerror = function() {
        showValidationError('无法加载Excel解析库，请使用CSV格式文件');
    };
    document.head.appendChild(script);
}

// 解析Excel文件
function parseExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const fileData = new Uint8Array(e.target.result);
        const workbook = XLSX.read(fileData, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        const importData = [];
        for (const row of jsonData) {
            // 尝试不同的表头名称
            const name = row['人员姓名'] || row['姓名'] || row['name'] || '';
            const count = parseInt(row['抽奖票数'] || row['票数'] || row['count'] || 0);
            
            if (name && !isNaN(count) && count >= 1 && count <= 100) {
                importData.push({ name, count });
            }
        }
        
        if (importData.length > 0) {
            importDataToTable(importData);
        } else {
            showValidationError('未找到有效的数据，请检查文件格式');
        }
    };
    reader.onerror = function() {
        showValidationError('文件读取失败');
    };
    reader.readAsArrayBuffer(file);
}

// 将数据导入到表格
function importDataToTable(data) {
    const tbody = participantsTable.querySelector('tbody');
    
    // 获取现有表格中的所有姓名
    const existingNames = new Set();
    const rows = tbody.querySelectorAll('tr');
    rows.forEach(row => {
        const nameInput = row.querySelector('.participant-name');
        if (nameInput) {
            const name = nameInput.value.trim();
            if (name) {
                existingNames.add(name);
            }
        }
    });
    
    // 过滤掉重复的姓名，只保留新数据
    const newData = data.filter(item => {
        const trimmedName = item.name.trim();
        return !existingNames.has(trimmedName);
    });
    
    // 添加新的不重复数据
    newData.forEach(item => {
        const lastRow = tbody.lastElementChild;
        addTableRow(lastRow, participantsTable);
        
        // 因为新行被插入到当前行前面，所以新行是lastRow的前一个兄弟节点
        const newRow = lastRow.previousElementSibling;
        const trimmedName = item.name.trim();
        newRow.querySelector('.participant-name').value = trimmedName;
        newRow.querySelector('.participant-count').value = item.count;
        
        // 更新删除按钮状态
        const deleteBtn = newRow.querySelector('.delete-row-btn');
        deleteBtn.disabled = tbody.rows.length <= 1;
    });
    
    // 清理表格末尾的空数据行（如果存在且不是唯一的行）
    const allRows = tbody.querySelectorAll('tr');
    if (allRows.length > 1) {
        const lastRow = allRows[allRows.length - 1];
        const lastRowName = lastRow.querySelector('.participant-name').value.trim();
        const lastRowCount = lastRow.querySelector('.participant-count').value.trim();
        
        if (!lastRowName && !lastRowCount) {
            lastRow.remove();
        }
    }
    
    // 更新删除按钮状态
    updateDeleteButtonState(participantsTable);
    
    // 显示导入成功提示
    if (newData.length > 0) {
        showSuccessMessage(`成功导入 ${newData.length} 条新数据`);
    } else {
        showSuccessMessage('没有新数据需要导入');
    }
}

// 显示成功消息
function showSuccessMessage(message) {
    // 创建临时提示元素
    const toast = document.createElement('div');
    toast.className = 'toast-message';
    toast.textContent = message;
    toast.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background-color: #4CAF50;
        color: white;
        padding: 15px 20px;
        border-radius: 5px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        z-index: 1000;
        animation: slideIn 0.3s ease-out;
    `;
    
    // 添加动画样式
    const style = document.createElement('style');
    style.textContent = `
        @keyframes slideIn {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }
    `;
    document.head.appendChild(style);
    
    document.body.appendChild(toast);
    
    // 3秒后移除提示
    setTimeout(() => {
        toast.style.animation = 'slideIn 0.3s ease-out reverse';
        setTimeout(() => {
            document.body.removeChild(toast);
            document.head.removeChild(style);
        }, 300);
    }, 3000);
}

// 重置抽奖
function resetLottery() {
    // 重置抽奖状态
    isRolling = false;
    clearInterval(rollingInterval);
    currentAward = null;
    awardParticipants = [];
    
    // 重置界面
    awardSelect.value = '';
    selectedAwardImage.style.display = 'none';
    selectedAwardName.textContent = '';
    awardQuota.textContent = '';
    remainingQuota.textContent = '';
    rollingNames.textContent = '点击开始抽奖';
    startBtn.disabled = true;
    stopBtn.disabled = true;
    resetAwardBtn.disabled = true;
    
    // 清空中奖记录
    recordsContainer.innerHTML = '';
    
    // 重置参与人员可用次数
    lotteryData.participants.forEach(participant => {
        participant.availableCount = participant.count;
    });
    
    // 重置奖项剩余名额
    lotteryData.awards.forEach(award => {
        award.remainingQuota = award.quota;
    });
    
    // 清空中奖记录
    lotteryData.winners = [];
}

// 重置当前奖项抽奖
function resetAwardLottery() {
    if (!currentAward) {
        alert('请先选择一个奖项');
        return;
    }
    
    if (confirm(`确定要重置${currentAward.name}的抽奖记录吗？\n这将清空该奖项的中奖记录，恢复默认名额，并允许已中奖人员重新参与。`)) {
        // 恢复奖项默认名额
        currentAward.remainingQuota = currentAward.quota;
        remainingQuota.textContent = currentAward.remainingQuota;
        
        // 记录需要恢复可用次数的人员
        const winnersToReset = lotteryData.winners.filter(winner => winner.awardName === currentAward.name);
        
        // 清空该奖项的中奖记录
        lotteryData.winners = lotteryData.winners.filter(winner => winner.awardName !== currentAward.name);
        
        // 恢复已中奖人员的可用次数
        winnersToReset.forEach(winner => {
            const participant = lotteryData.participants.find(p => p.name === winner.name);
            if (participant) {
                participant.availableCount++;
            }
        });
        
        // 重新渲染中奖记录
        renderWinnerRecords();
        
        // 更新开始抽奖按钮状态
        startBtn.disabled = false;
        
        alert(`${currentAward.name}的抽奖记录已重置成功！`);
    }
}

// 重新渲染中奖记录
function renderWinnerRecords() {
    // 清空记录容器
    recordsContainer.innerHTML = '';
    
    // 按时间倒序添加记录
    const sortedWinners = [...lotteryData.winners].sort((a, b) => b.time - a.time);
    
    sortedWinners.forEach(winner => {
        const recordItem = document.createElement('div');
        recordItem.className = 'record-item';
        
        const timeStr = formatDateTime(winner.time);
        
        recordItem.innerHTML = `
            <span class="record-name">${winner.name}</span>
            <span class="record-award">${winner.awardName}</span>
            <span class="record-time">${timeStr}</span>
        `;
        
        recordsContainer.appendChild(recordItem);
    });
}

// 导出Excel功能
function exportToExcel() {
    // 检查是否有中奖记录
    if (lotteryData.winners.length === 0) {
        alert('暂无中奖记录可导出');
        return;
    }
    
    // 创建CSV内容
    let csvContent = "中奖人员姓名,中奖奖项,中奖时间\n";
    
    // 添加中奖记录数据
    lotteryData.winners.forEach(winner => {
        const timeStr = formatDateTime(winner.time);
        // 确保CSV格式正确，处理包含逗号的字段
        csvContent += `"${winner.name}","${winner.awardName}","${timeStr}"\n`;
    });
    
    // 创建Blob对象
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    
    // 创建下载链接
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    // 设置下载属性
    link.setAttribute('href', url);
    link.setAttribute('download', `${lotteryData.activityName}_中奖记录.csv`);
    link.style.visibility = 'hidden';
    
    // 添加到DOM并触发点击
    document.body.appendChild(link);
    link.click();
    
    // 清理
    document.body.removeChild(link);
}