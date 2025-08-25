document.addEventListener('DOMContentLoaded', function() {
    // 全局变量
    let pages = [];
    let currentPageIndex = 0;
    let colorPickerTarget = null;
    let colorPickerModal = new bootstrap.Modal(document.getElementById('colorPickerModal'));
    
    // DOM元素
    const notionContent = document.getElementById('notionContent');
    const generatePreviewBtn = document.getElementById('generatePreview');
    const imageContainer = document.getElementById('imageContainer');
    const prevPageBtn = document.getElementById('prevPage');
    const nextPageBtn = document.getElementById('nextPage');
    const currentPageSpan = document.getElementById('currentPage');
    const totalPagesSpan = document.getElementById('totalPages');
    const downloadCurrentBtn = document.getElementById('downloadCurrent');
    const downloadAllBtn = document.getElementById('downloadAll');
    
    // 初始化事件监听器
    initEventListeners();
    
    // 初始化事件监听器
    function initEventListeners() {
        // 初始化默认尺寸
        setImageSizeByRatio('3:4');
        
        // 生成预览
        generatePreviewBtn.addEventListener('click', generatePreview);
        
        // 页面导航
        prevPageBtn.addEventListener('click', showPreviousPage);
        nextPageBtn.addEventListener('click', showNextPage);
        
        // 下载按钮
        downloadCurrentBtn.addEventListener('click', downloadCurrentPage);
        downloadAllBtn.addEventListener('click', downloadAllPages);
        
        // 尺寸预设
        document.querySelectorAll('.size-preset').forEach(button => {
            button.addEventListener('click', function() {
                document.querySelectorAll('.size-preset').forEach(btn => btn.classList.remove('active'));
                this.classList.add('active');
                
                const ratio = this.dataset.ratio;
                setImageSizeByRatio(ratio);
            });
        });
        
        // 字号预设
        document.querySelectorAll('.font-size-preset').forEach(button => {
            button.addEventListener('click', function() {
                document.querySelectorAll('.font-size-preset').forEach(btn => btn.classList.remove('active'));
                this.classList.add('active');
                
                const size = this.dataset.size;
                setFontSizePreset(size);
            });
        });
        
        // 配色方案预设
        document.querySelectorAll('.color-preset').forEach(button => {
            button.addEventListener('click', function() {
                document.querySelectorAll('.color-preset').forEach(btn => btn.classList.remove('active'));
                this.classList.add('active');
                
                const bg = this.dataset.bg;
                const title = this.dataset.title;
                const text = this.dataset.text;
                const titleBg = this.dataset.titleBg;
                
                setColorSchemeByData(bg, title, text, titleBg);
            });
        });
        
        // 颜色选择器
        document.querySelectorAll('.color-picker').forEach(button => {
            button.addEventListener('click', function() {
                colorPickerTarget = this.dataset.target;
                const currentColor = document.getElementById(colorPickerTarget).value;
                document.getElementById('colorPicker').value = currentColor;
                colorPickerModal.show();
            });
        });
        
        // 应用颜色
        document.getElementById('applyColor').addEventListener('click', function() {
            const selectedColor = document.getElementById('colorPicker').value;
            document.getElementById(colorPickerTarget).value = selectedColor;
            colorPickerModal.hide();
            
            // 如果已经生成了预览，则更新预览
            if (pages.length > 0) {
                updatePreview();
            }
        });
        
        // 监听字体大小变化
        document.getElementById('bodyFontSize').addEventListener('input', function() {
            updateHeadingFontSizes();
            
            // 如果已经生成了预览，则更新预览
            if (pages.length > 0) {
                updatePreview();
            }
        });
        
        // 监听其他设置变化
        const settingsInputs = [
            'imageWidth', 'imageHeight', 'horizontalPadding', 'verticalPadding',
            'lineHeight', 'fontFamily', 'backgroundColor', 'headingColor', 'bodyColor'
        ];
        
        settingsInputs.forEach(id => {
            document.getElementById(id).addEventListener('change', function() {
                if (pages.length > 0) {
                    updatePreview();
                }
            });
        });
        
        // 监听对齐方式变化
        document.querySelectorAll('input[name="textAlign"]').forEach(radio => {
            radio.addEventListener('change', function() {
                if (pages.length > 0) {
                    updatePreview();
                }
            });
        });
    }
    

    
    // 生成预览
    function generatePreview() {
        const content = notionContent.value.trim();
        
        if (!content) {
            alert('请输入Notion内容');
            return;
        }
        
        // 解析Notion内容
        const parsedBlocks = parseNotionContent(content);
        
        // 分页处理
        pages = paginateContent(parsedBlocks);
        
        // 更新页面指示器
        currentPageIndex = 0;
        updatePageIndicator();
        
        // 显示第一页
        renderPage(currentPageIndex);
    }
    
    // 解析Notion内容
    function parseNotionContent(content) {
        // 按行分割内容
        const lines = content.split('\n');
        const blocks = [];
        
        let currentBlock = null;
        
        lines.forEach(line => {
            // 跳过空行
            if (line.trim() === '') {
                if (currentBlock) {
                    blocks.push(currentBlock);
                    currentBlock = null;
                }
                return;
            }
            
            // 检测块类型
            if (line.startsWith('# ')) {
                // Heading 1
                if (currentBlock) {
                    blocks.push(currentBlock);
                }
                currentBlock = { type: 'heading_1', content: line.substring(2).trim() };
                blocks.push(currentBlock);
                currentBlock = null;
            } else if (line.startsWith('## ')) {
                // Heading 2
                if (currentBlock) {
                    blocks.push(currentBlock);
                }
                currentBlock = { type: 'heading_2', content: line.substring(3).trim() };
                blocks.push(currentBlock);
                currentBlock = null;
            } else if (line.startsWith('### ')) {
                // Heading 3
                if (currentBlock) {
                    blocks.push(currentBlock);
                }
                currentBlock = { type: 'heading_3', content: line.substring(4).trim() };
                blocks.push(currentBlock);
                currentBlock = null;
            } else if (line.startsWith('- ')) {
                // Unordered list
                if (!currentBlock || currentBlock.type !== 'unordered_list') {
                    if (currentBlock) {
                        blocks.push(currentBlock);
                    }
                    currentBlock = { type: 'unordered_list', items: [] };
                }
                currentBlock.items.push(line.substring(2).trim());
            } else if (line.match(/^\d+\.\s/)) {
                // Ordered list
                if (!currentBlock || currentBlock.type !== 'ordered_list') {
                    if (currentBlock) {
                        blocks.push(currentBlock);
                    }
                    currentBlock = { type: 'ordered_list', items: [] };
                }
                currentBlock.items.push(line.substring(line.indexOf('.') + 1).trim());
            } else if (line.startsWith('> ')) {
                // Quote
                if (!currentBlock || currentBlock.type !== 'quote') {
                    if (currentBlock) {
                        blocks.push(currentBlock);
                    }
                    currentBlock = { type: 'quote', content: line.substring(2).trim() };
                } else {
                    currentBlock.content += '\n' + line.substring(2).trim();
                }
            } else if (line.startsWith('```')) {
                // Code block start or end
                if (!currentBlock || currentBlock.type !== 'code') {
                    if (currentBlock) {
                        blocks.push(currentBlock);
                    }
                    currentBlock = { type: 'code', content: '' };
                } else {
                    blocks.push(currentBlock);
                    currentBlock = null;
                }
            } else if (currentBlock && currentBlock.type === 'code') {
                // Code block content
                if (currentBlock.content) {
                    currentBlock.content += '\n' + line;
                } else {
                    currentBlock.content = line;
                }
            } else {
                // Paragraph
                if (!currentBlock || currentBlock.type !== 'paragraph') {
                    if (currentBlock) {
                        blocks.push(currentBlock);
                    }
                    currentBlock = { type: 'paragraph', content: line.trim() };
                } else {
                    currentBlock.content += '\n' + line.trim();
                }
            }
        });
        
        // 添加最后一个块
        if (currentBlock) {
            blocks.push(currentBlock);
        }
        
        // 处理富文本样式
        blocks.forEach(block => {
            if (block.content) {
                block.content = processRichTextStyles(block.content);
            } else if (block.items) {
                block.items = block.items.map(item => processRichTextStyles(item));
            }
        });
        
        return blocks;
    }
    
    // 处理富文本样式
    function processRichTextStyles(text) {
        // 先处理代码块内的内容，避免被其他规则影响
        const codeBlocks = [];
        text = text.replace(/`([^`]+)`/g, (match, code) => {
            const index = codeBlocks.length;
            codeBlocks.push(`<code>${code}</code>`);
            return `__CODE_BLOCK_${index}__`;
        });
        
        // 处理加粗 **text** 或 __text__ (非贪婪匹配)
        text = text.replace(/\*\*([^\*]+?)\*\*/g, '<span class="bold">$1</span>');
        text = text.replace(/__([^_]+?)__/g, '<span class="bold">$1</span>');
        
        // 处理斜体 *text* 或 _text_ (确保不与加粗冲突)
        text = text.replace(/(?<!\*)\*([^\*\n]+?)\*(?!\*)/g, '<span class="italic">$1</span>');
        text = text.replace(/(?<!_)_([^_\n]+?)_(?!_)/g, '<span class="italic">$1</span>');
        
        // 处理删除线 ~~text~~
        text = text.replace(/~~([^~]+?)~~/g, '<span class="strikethrough">$1</span>');
        
        // 处理下划线 ++text++
        text = text.replace(/\+\+([^\+]+?)\+\+/g, '<span class="underline">$1</span>');
        
        // 处理链接 [text](url) - 转换为纯文本
        text = text.replace(/\[([^\]]+?)\]\([^\)]+?\)/g, '$1');
        
        // 恢复代码块
        codeBlocks.forEach((code, index) => {
            text = text.replace(`__CODE_BLOCK_${index}__`, code);
        });
        
        return text;
    }
    
    // 检测字符类型
    function getCharacterType(char) {
        const code = char.charCodeAt(0);
        
        // Emoji 检测（基本范围）
        if ((code >= 0x1F600 && code <= 0x1F64F) || // 表情符号
            (code >= 0x1F300 && code <= 0x1F5FF) || // 杂项符号
            (code >= 0x1F680 && code <= 0x1F6FF) || // 交通和地图符号
            (code >= 0x2600 && code <= 0x26FF) ||   // 杂项符号
            (code >= 0x2700 && code <= 0x27BF)) {   // 装饰符号
            return 'emoji';
        }
        
        // 中文字符检测
        if ((code >= 0x4E00 && code <= 0x9FFF) ||   // CJK统一汉字
            (code >= 0x3400 && code <= 0x4DBF) ||   // CJK扩展A
            (code >= 0x20000 && code <= 0x2A6DF) || // CJK扩展B
            (code >= 0x3000 && code <= 0x303F) ||   // CJK符号和标点
            (code >= 0xFF00 && code <= 0xFFEF)) {   // 全角ASCII
            return 'chinese';
        }
        
        // 制表符
        if (char === '\t') {
            return 'tab';
        }
        
        // 换行符
        if (char === '\n') {
            return 'newline';
        }
        
        // 其他字符视为英文/半角
        return 'english';
    }
    
    // 计算字符的实际宽度
    function getCharacterWidth(char, fontSize, fontFamily) {
        const type = getCharacterType(char);
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        ctx.font = `${fontSize}px ${getFontFamilyValue(fontFamily)}`;
        
        switch (type) {
            case 'emoji':
                // Emoji通常占用2个字符宽度
                return ctx.measureText('中').width * 2;
            case 'chinese':
                // 中文字符使用实际测量
                return ctx.measureText(char).width;
            case 'tab':
                // 制表符按4个空格计算
                return ctx.measureText('    ').width;
            case 'newline':
                // 换行符不占宽度
                return 0;
            default:
                // 英文字符使用实际测量
                return ctx.measureText(char).width;
        }
    }
    
    // 精确测量文本尺寸的辅助函数
    function measureText(text, fontSize, fontFamily) {
        // 创建临时canvas用于测量
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        ctx.font = `${fontSize}px ${getFontFamilyValue(fontFamily)}`;
        
        // 处理特殊字符和换行符
        const lines = text.split('\n');
        let maxWidth = 0;
        let totalHeight = 0;
        
        lines.forEach(line => {
            if (line.trim() === '') {
                // 空行也占用高度
                totalHeight += fontSize * 1.2; // 默认行高
            } else {
                // 逐字符精确测量宽度
                let lineWidth = 0;
                for (let i = 0; i < line.length; i++) {
                    lineWidth += getCharacterWidth(line[i], fontSize, fontFamily);
                }
                maxWidth = Math.max(maxWidth, lineWidth);
                totalHeight += fontSize * 1.2;
            }
        });
        
        return {
            width: maxWidth,
            height: totalHeight,
            lines: lines.length
        };
    }
    
    // 计算文本在指定宽度内需要的行数
    function calculateTextLines(text, maxWidth, fontSize, fontFamily) {
        if (!text || text.trim() === '') {
            return 1;
        }
        
        // 处理换行符
        const paragraphs = text.split('\n');
        let totalLines = 0;
        
        paragraphs.forEach(paragraph => {
            if (paragraph.trim() === '') {
                totalLines += 1; // 空行
                return;
            }
            
            // 逐字符精确测量，处理自动换行
            let currentWidth = 0;
            let lineCount = 1;
            
            for (let i = 0; i < paragraph.length; i++) {
                const char = paragraph[i];
                const charWidth = getCharacterWidth(char, fontSize, fontFamily);
                
                // 检查是否需要换行
                if (currentWidth + charWidth > maxWidth) {
                    // 如果是空格，可以在此处换行
                    if (char === ' ') {
                        lineCount++;
                        currentWidth = 0;
                        continue;
                    }
                    
                    // 查找最近的空格进行换行
                    let canBreak = false;
                    for (let j = i - 1; j >= 0; j--) {
                        if (paragraph[j] === ' ') {
                            canBreak = true;
                            break;
                        }
                        // 如果回溯太远，强制换行
                        if (i - j > 20) {
                            break;
                        }
                    }
                    
                    lineCount++;
                    currentWidth = charWidth;
                } else {
                    currentWidth += charWidth;
                }
            }
            
            totalLines += Math.max(1, lineCount);
        });
        
        return totalLines;
    }
    
    // 分页处理
    function paginateContent(blocks) {
        const imageWidth = parseInt(document.getElementById('imageWidth').value);
        const imageHeight = parseInt(document.getElementById('imageHeight').value);
        const horizontalPaddingPercent = parseFloat(document.getElementById('horizontalPadding').value);
        const verticalPaddingPercent = parseFloat(document.getElementById('verticalPadding').value);
        const bodyFontSize = parseInt(document.getElementById('bodyFontSize').value);
        const lineHeight = parseFloat(document.getElementById('lineHeight').value);
        const fontFamily = document.getElementById('fontFamily').value;
        
        // 将百分比转换为像素值
        const horizontalPadding = Math.round(imageWidth * horizontalPaddingPercent / 100);
        const verticalPadding = Math.round(imageHeight * verticalPaddingPercent / 100);
        
        // 计算可用高度和宽度
        const availableHeight = imageHeight - (verticalPadding * 2);
        const availableWidth = imageWidth - (horizontalPadding * 2);
        
        // 获取字体渲染补偿系数
        function getFontRenderingCompensation(fontFamily) {
            const fontName = getFontFamilyValue(fontFamily).toLowerCase();
            
            // 不同字体的渲染差异补偿
            const compensationMap = {
                'simhei': 1.1,      // 黑体通常渲染较宽
                'simsun': 1.05,     // 宋体
                'microsoftyahei': 1.08, // 微软雅黑
                'arial': 0.95,      // Arial较窄
                'helvetica': 0.95,  // Helvetica较窄
                'times': 1.0,       // Times基准
                'courier': 1.2      // 等宽字体较宽
            };
            
            for (let font in compensationMap) {
                if (fontName.includes(font)) {
                    return compensationMap[font];
                }
            }
            
            return 1.0; // 默认无补偿
        }
        
        const fontCompensation = getFontRenderingCompensation(fontFamily);
        
        // 更精确地估算每个块的高度
        const estimatedHeights = blocks.map((block, index) => {
            let height = 0;
            const marginBottom = bodyFontSize * 0.5; // 块之间的间距
            
            // 添加基线偏移补偿（PIL库渲染特性）
            const baselineCompensation = bodyFontSize * 0.1;
            
            try {
                switch (block.type) {
                    case 'heading_1':
                        const h1FontSize = Math.round(bodyFontSize * 1.8);
                        const h1Content = block.content ? block.content.replace(/<[^>]*>/g, '') : '';
                        if (h1Content.trim()) {
                            const h1Lines = calculateTextLines(h1Content, availableWidth * fontCompensation, h1FontSize, fontFamily);
                            height = h1Lines * h1FontSize * lineHeight + marginBottom + baselineCompensation;
                        } else {
                            height = h1FontSize * lineHeight + marginBottom;
                        }
                        break;
                    case 'heading_2':
                        const h2FontSize = Math.round(bodyFontSize * 1.5);
                        const h2Content = block.content ? block.content.replace(/<[^>]*>/g, '') : '';
                        if (h2Content.trim()) {
                            const h2Lines = calculateTextLines(h2Content, availableWidth * fontCompensation, h2FontSize, fontFamily);
                            height = h2Lines * h2FontSize * lineHeight + marginBottom + baselineCompensation;
                        } else {
                            height = h2FontSize * lineHeight + marginBottom;
                        }
                        break;
                    case 'heading_3':
                        const h3FontSize = Math.round(bodyFontSize * 1.25);
                        const h3Content = block.content ? block.content.replace(/<[^>]*>/g, '') : '';
                        if (h3Content.trim()) {
                            const h3Lines = calculateTextLines(h3Content, availableWidth * fontCompensation, h3FontSize, fontFamily);
                            height = h3Lines * h3FontSize * lineHeight + marginBottom + baselineCompensation;
                        } else {
                            height = h3FontSize * lineHeight + marginBottom;
                        }
                        break;
                    case 'paragraph':
                        // 精确的文本高度计算
                        const textContent = block.content ? block.content.replace(/<[^>]*>/g, '') : '';
                        if (textContent.trim()) {
                            const lines = calculateTextLines(textContent, availableWidth * fontCompensation, bodyFontSize, fontFamily);
                            height = lines * bodyFontSize * lineHeight + marginBottom + baselineCompensation;
                        } else {
                            // 空段落仍占一行高度
                            height = bodyFontSize * lineHeight + marginBottom;
                        }
                        break;
                    case 'unordered_list':
                    case 'ordered_list':
                        // 计算列表项的总高度
                        let listHeight = 0;
                        if (block.items && block.items.length > 0) {
                            block.items.forEach(item => {
                                const itemContent = item.replace(/<[^>]*>/g, '');
                                if (itemContent.trim()) {
                                    const itemLines = calculateTextLines(itemContent, (availableWidth - 30) * fontCompensation, bodyFontSize, fontFamily);
                                    listHeight += itemLines * bodyFontSize * lineHeight + baselineCompensation;
                                } else {
                                    listHeight += bodyFontSize * lineHeight;
                                }
                            });
                        } else {
                            listHeight = bodyFontSize * lineHeight;
                        }
                        height = listHeight + marginBottom;
                        break;
                    case 'quote':
                        const quoteContent = block.content ? block.content.replace(/<[^>]*>/g, '') : '';
                        if (quoteContent.trim()) {
                            const quoteLines = calculateTextLines(quoteContent, (availableWidth - 20) * fontCompensation, bodyFontSize, fontFamily);
                            height = quoteLines * bodyFontSize * lineHeight + marginBottom + 20 + baselineCompensation;
                        } else {
                            height = bodyFontSize * lineHeight + marginBottom + 20;
                        }
                        break;
                    case 'code':
                        // 代码块按实际换行符计算行数
                        const codeContent = block.content || '';
                        const codeLines = Math.max(1, codeContent.split('\n').length);
                        // 代码块使用等宽字体，需要特殊处理
                        height = codeLines * bodyFontSize * lineHeight * 1.1 + marginBottom + 40 + baselineCompensation;
                        break;
                    default:
                        // 处理未知类型
                        console.warn(`未知块类型: ${block.type}`);
                        height = bodyFontSize * lineHeight + marginBottom;
                        break;
                }
            } catch (error) {
                console.error(`计算块 ${index} (${block.type}) 高度时出错:`, error);
                height = bodyFontSize * lineHeight + marginBottom;
            }
            
            // 确保最小高度
            const minHeight = bodyFontSize * lineHeight * 0.5;
            return Math.max(height, minHeight);
        });
        
        // 边界条件检查
        if (!blocks || blocks.length === 0) {
            return [[]];
        }
        
        if (availableHeight <= 0 || availableWidth <= 0) {
            console.warn('可用空间不足，无法进行分页');
            return [[]];
        }
        
        // 精确的分页算法 - 处理内容截断问题
        const result = [];
        let currentPage = [];
        let currentHeight = 0;
        let i = 0;
        
        // 防止无限循环的安全计数器
        let safetyCounter = 0;
        const maxIterations = blocks.length * 10;
        
        while (i < blocks.length && safetyCounter < maxIterations) {
            safetyCounter++;
            
            const block = blocks[i];
            if (!block) {
                console.warn(`块索引 ${i} 为空，跳过`);
                i++;
                continue;
            }
            
            const blockHeight = estimatedHeights[i] || 0;
            
            // 检查块高度是否合理
            if (blockHeight <= 0) {
                console.warn(`块 ${i} 高度为 ${blockHeight}，使用默认高度`);
                blockHeight = bodyFontSize * lineHeight;
            }
            
            // 检查当前块是否能完整放入当前页
            if (currentHeight + blockHeight > availableHeight) {
                // 内容会被截断，需要分页处理
                
                if (currentPage.length === 0) {
                    // 当前页为空但内容仍然太大，强制添加并警告
                    console.warn(`块 ${i} (${block.type}) 高度 ${blockHeight} 超过页面可用高度 ${availableHeight}，强制添加到页面`);
                    currentPage.push(block);
                    result.push([...currentPage]);
                    currentPage = [];
                    currentHeight = 0;
                } else {
                    // 判断被截断的是标题还是正文
                    if (['heading_1', 'heading_2', 'heading_3'].includes(block.type)) {
                        // 被截断的是标题，直接将标题和后续内容放到下一页
                        result.push([...currentPage]);
                        currentPage = [block];
                        currentHeight = blockHeight;
                    } else {
                        // 被截断的是正文，需要回溯到最近的标题
                        let lastHeadingIndex = -1;
                        
                        // 从当前页的最后一个元素开始向前查找标题
                        for (let j = currentPage.length - 1; j >= 0; j--) {
                            if (currentPage[j] && ['heading_1', 'heading_2', 'heading_3'].includes(currentPage[j].type)) {
                                lastHeadingIndex = j;
                                break;
                            }
                        }
                        
                        if (lastHeadingIndex !== -1) {
                            // 记录要移动的块在原始数组中的起始索引
                            const moveStartIndex = i - (currentPage.length - lastHeadingIndex);
                            
                            // 找到了标题，将标题及其后的内容移到新页
                            const blocksToMove = currentPage.splice(lastHeadingIndex);
                            
                            // 保存当前页（标题之前的内容）
                            if (currentPage.length > 0) {
                                result.push([...currentPage]);
                            }
                            
                            // 开始新页，包含移动的标题和内容，以及当前被截断的块
                            currentPage = [...blocksToMove, block];
                            
                            // 重新计算新页的高度
                            currentHeight = 0;
                            // 计算移动的块的高度
                            for (let k = 0; k < blocksToMove.length; k++) {
                                const moveBlockIndex = moveStartIndex + k;
                                if (moveBlockIndex >= 0 && moveBlockIndex < estimatedHeights.length) {
                                    currentHeight += estimatedHeights[moveBlockIndex] || 0;
                                }
                            }
                            currentHeight += blockHeight;
                        } else {
                            // 没有找到标题，直接分页
                            result.push([...currentPage]);
                            currentPage = [block];
                            currentHeight = blockHeight;
                        }
                    }
                }
            } else {
                // 当前块可以完整放入当前页
                currentPage.push(block);
                currentHeight += blockHeight;
            }
            
            i++;
        }
        
        // 检查是否因为安全计数器而退出循环
        if (safetyCounter >= maxIterations) {
            console.error('分页算法可能陷入无限循环，强制退出');
        }
        
        // 添加最后一页
        if (currentPage.length > 0) {
            result.push([...currentPage]);
        }
        
        // 确保至少有一页
        if (result.length === 0) {
            result.push([]);
        }
        
        // 验证结果
        console.log(`分页完成：共 ${result.length} 页，原始块数：${blocks.length}`);
        
        return result;
    }
    
    // 渲染页面
    function renderPage(pageIndex) {
        if (pageIndex < 0 || pageIndex >= pages.length) {
            return;
        }
        
        const page = pages[pageIndex];
        const imageWidth = parseInt(document.getElementById('imageWidth').value);
        const imageHeight = parseInt(document.getElementById('imageHeight').value);
        const horizontalPaddingPercent = parseFloat(document.getElementById('horizontalPadding').value);
        const verticalPaddingPercent = parseFloat(document.getElementById('verticalPadding').value);
        
        // 将百分比转换为像素值
        const horizontalPadding = Math.round(imageWidth * horizontalPaddingPercent / 100);
        const verticalPadding = Math.round(imageHeight * verticalPaddingPercent / 100);
        const fontFamily = document.getElementById('fontFamily').value;
        const backgroundColor = document.getElementById('backgroundColor').value;
        const headingColor = document.getElementById('headingColor').value;
        const bodyColor = document.getElementById('bodyColor').value;
        const bodyFontSize = parseInt(document.getElementById('bodyFontSize').value);
        const h1FontSize = parseInt(document.getElementById('h1FontSize').value);
        const h2FontSize = parseInt(document.getElementById('h2FontSize').value);
        const h3FontSize = parseInt(document.getElementById('h3FontSize').value);
        const lineHeight = parseFloat(document.getElementById('lineHeight').value);
        const textAlign = document.querySelector('input[name="textAlign"]:checked').value;
        
        // 创建页面容器
        const pageElement = document.createElement('div');
        pageElement.className = 'notion-page';
        pageElement.style.width = `${imageWidth}px`;
        pageElement.style.height = `${imageHeight}px`;
        pageElement.style.backgroundColor = backgroundColor;
        pageElement.style.padding = `${verticalPadding}px ${horizontalPadding}px`;
        pageElement.style.fontFamily = getFontFamilyValue(fontFamily);
        pageElement.style.color = bodyColor;
        pageElement.style.fontSize = `${bodyFontSize}px`;
        pageElement.style.lineHeight = lineHeight;
        pageElement.style.textAlign = textAlign;
        pageElement.style.overflow = 'hidden';
        pageElement.style.position = 'relative';
        
        // 创建内容容器
        const contentElement = document.createElement('div');
        contentElement.className = 'notion-content';
        
        // 渲染块
        page.forEach(block => {
            const blockElement = renderBlock(block, {
                headingColor,
                bodyColor,
                h1FontSize,
                h2FontSize,
                h3FontSize,
                bodyFontSize
            });
            
            if (blockElement) {
                contentElement.appendChild(blockElement);
            }
        });
        
        pageElement.appendChild(contentElement);
        
        // 更新预览区域
        imageContainer.innerHTML = '';
        imageContainer.appendChild(pageElement);
        
        // 更新导航按钮状态
        prevPageBtn.disabled = pageIndex === 0;
        nextPageBtn.disabled = pageIndex === pages.length - 1;
    }
    
    // 渲染块
    function renderBlock(block, styles) {
        const { headingColor, bodyColor, h1FontSize, h2FontSize, h3FontSize, bodyFontSize } = styles;
        
        let element;
        
        switch (block.type) {
            case 'heading_1':
                element = document.createElement('h1');
                element.innerHTML = block.content;
                element.style.fontSize = `${h1FontSize}px`;
                element.style.color = headingColor;
                break;
                
            case 'heading_2':
                element = document.createElement('h2');
                element.innerHTML = block.content;
                element.style.fontSize = `${h2FontSize}px`;
                element.style.color = headingColor;
                break;
                
            case 'heading_3':
                element = document.createElement('h3');
                element.innerHTML = block.content;
                element.style.fontSize = `${h3FontSize}px`;
                element.style.color = headingColor;
                break;
                
            case 'paragraph':
                element = document.createElement('p');
                element.innerHTML = block.content;
                element.style.fontSize = `${bodyFontSize}px`;
                element.style.color = bodyColor;
                break;
                
            case 'unordered_list':
                element = document.createElement('ul');
                block.items.forEach(item => {
                    const li = document.createElement('li');
                    li.innerHTML = item;
                    li.style.fontSize = `${bodyFontSize}px`;
                    li.style.color = bodyColor;
                    element.appendChild(li);
                });
                break;
                
            case 'ordered_list':
                element = document.createElement('ol');
                block.items.forEach(item => {
                    const li = document.createElement('li');
                    li.innerHTML = item;
                    li.style.fontSize = `${bodyFontSize}px`;
                    li.style.color = bodyColor;
                    element.appendChild(li);
                });
                break;
                
            case 'quote':
                element = document.createElement('blockquote');
                element.innerHTML = block.content;
                element.style.fontSize = `${bodyFontSize}px`;
                element.style.color = bodyColor;
                break;
                
            case 'code':
                element = document.createElement('pre');
                const code = document.createElement('code');
                code.textContent = block.content;
                code.style.fontSize = `${bodyFontSize}px`;
                code.style.color = bodyColor;
                element.appendChild(code);
                break;
                
            default:
                return null;
        }
        
        return element;
    }
    
    // 更新页面指示器
    function updatePageIndicator() {
        currentPageSpan.textContent = currentPageIndex + 1;
        totalPagesSpan.textContent = pages.length;
    }
    
    // 显示上一页
    function showPreviousPage() {
        if (currentPageIndex > 0) {
            currentPageIndex--;
            renderPage(currentPageIndex);
            updatePageIndicator();
            setTimeout(autoScalePreview, 100);
        }
    }
    
    // 显示下一页
    function showNextPage() {
        if (currentPageIndex < pages.length - 1) {
            currentPageIndex++;
            renderPage(currentPageIndex);
            updatePageIndicator();
            setTimeout(autoScalePreview, 100);
        }
    }
    
    // 自动缩放预览容器
    function autoScalePreview() {
        const imageContainer = document.getElementById('imageContainer');
        const imageWidth = parseInt(document.getElementById('imageWidth').value);
        const imageHeight = parseInt(document.getElementById('imageHeight').value);
        
        // 获取容器的可用空间
        const containerRect = imageContainer.getBoundingClientRect();
        const maxWidth = containerRect.width - 40; // 留一些边距
        const maxHeight = window.innerHeight - 200; // 留一些空间给其他元素
        
        // 计算缩放比例
        const scaleX = maxWidth / imageWidth;
        const scaleY = maxHeight / imageHeight;
        const scale = Math.min(scaleX, scaleY, 1); // 不放大，只缩小
        
        // 应用缩放
        const previewElement = imageContainer.querySelector('.notion-page');
        if (previewElement) {
            previewElement.style.transform = `scale(${scale})`;
            previewElement.style.transformOrigin = 'top left';
            
            // 调整容器高度以适应缩放后的内容
            imageContainer.style.height = `${imageHeight * scale + 40}px`;
            imageContainer.style.overflow = 'visible';
        }
    }
    
    // 更新预览
    function updatePreview() {
        // 重新解析原始内容
        const content = notionContent.value.trim();
        
        if (!content) {
            return;
        }
        
        // 解析Notion内容
        const parsedBlocks = parseNotionContent(content);
        
        // 重新分页
        pages = paginateContent(parsedBlocks);
        
        // 更新页面指示器
        updatePageIndicator();
        
        // 如果当前页索引超出范围，则重置为最后一页
        if (currentPageIndex >= pages.length) {
            currentPageIndex = pages.length - 1;
        }
        
        // 显示当前页
        renderPage(currentPageIndex);
        
        // 延迟执行自动缩放，确保DOM已更新
        setTimeout(autoScalePreview, 100);
    }
    
    // 下载当前页面
    function downloadCurrentPage() {
        if (pages.length === 0) {
            alert('请先生成预览');
            return;
        }
        
        const format = document.querySelector('input[name="downloadFormat"]:checked').value;
        const pageElement = document.querySelector('.notion-page');
        
        html2canvas(pageElement, {
            scale: 2, // 提高导出图片质量
            useCORS: true,
            backgroundColor: null
        }).then(canvas => {
            const link = document.createElement('a');
            link.download = `notion-image-${currentPageIndex + 1}.${format}`;
            
            if (format === 'png') {
                link.href = canvas.toDataURL('image/png');
            } else {
                link.href = canvas.toDataURL('image/jpeg', 0.9);
            }
            
            link.click();
        });
    }
    
    // 下载所有页面
    function downloadAllPages() {
        if (pages.length === 0) {
            alert('请先生成预览');
            return;
        }
        
        const format = document.querySelector('input[name="downloadFormat"]:checked').value;
        const originalPageIndex = currentPageIndex;
        
        // 创建一个zip文件
        alert('正在准备下载所有页面，请稍候...');
        
        // 逐个下载每一页
        const downloadPage = (index) => {
            if (index >= pages.length) {
                // 恢复原始页面
                currentPageIndex = originalPageIndex;
                renderPage(currentPageIndex);
                updatePageIndicator();
                alert('所有页面已下载完成');
                return;
            }
            
            // 显示当前页
            currentPageIndex = index;
            renderPage(currentPageIndex);
            updatePageIndicator();
            
            // 延迟一下，确保页面已渲染
            setTimeout(() => {
                const pageElement = document.querySelector('.notion-page');
                
                html2canvas(pageElement, {
                    scale: 2,
                    useCORS: true,
                    backgroundColor: null
                }).then(canvas => {
                    const link = document.createElement('a');
                    link.download = `notion-image-${index + 1}.${format}`;
                    
                    if (format === 'png') {
                        link.href = canvas.toDataURL('image/png');
                    } else {
                        link.href = canvas.toDataURL('image/jpeg', 0.9);
                    }
                    
                    link.click();
                    
                    // 下载下一页
                    setTimeout(() => {
                        downloadPage(index + 1);
                    }, 500);
                });
            }, 100);
        };
        
        // 开始下载第一页
        downloadPage(0);
    }
    
    // 根据比例设置图片尺寸
    function setImageSizeByRatio(ratio) {
        const [width, height] = ratio.split(':').map(Number);
        const scale = 200; // 缩放因子
        
        document.getElementById('imageWidth').value = width * scale;
        document.getElementById('imageHeight').value = height * scale;
        
        // 如果已经生成了预览，则更新预览
        if (pages.length > 0) {
            updatePreview();
        }
    }
    
    // 设置字号预设
    function setFontSizePreset(size) {
        let bodySize;
        
        switch (size) {
            case 'small':
                bodySize = 14;
                break;
            case 'medium':
                bodySize = 16;
                break;
            case 'large':
                bodySize = 18;
                break;
            default:
                bodySize = 16;
        }
        
        document.getElementById('bodyFontSize').value = bodySize;
        updateHeadingFontSizes();
        
        // 如果已经生成了预览，则更新预览
        if (pages.length > 0) {
            updatePreview();
        }
    }
    
    // 更新标题字号
    function updateHeadingFontSizes() {
        const bodySize = parseInt(document.getElementById('bodyFontSize').value);
        
        document.getElementById('h3FontSize').value = Math.round(bodySize * 1.25);
        document.getElementById('h2FontSize').value = Math.round(bodySize * 1.5);
        document.getElementById('h1FontSize').value = Math.round(bodySize * 1.8);
    }
    
    // 设置配色方案（通过数据）
    function setColorSchemeByData(bg, title, text, titleBg) {
        document.getElementById('backgroundColor').value = bg;
        document.getElementById('headingColor').value = title;
        document.getElementById('bodyColor').value = text;
        
        // 如果已经生成了预览，则更新预览
        if (pages.length > 0) {
            updatePreview();
        }
    }
    
    // 设置配色方案（保留兼容性）
    function setColorScheme(scheme) {
        let background, heading, body;
        
        switch (scheme) {
            case 'simple':
                background = '#F9F9F9';
                heading = '#222222';
                body = '#444444';
                break;
            case 'dopamine':
                background = '#FFF0F3';
                heading = '#FFD100';
                body = '#42A5F5';
                break;
            case 'night':
                background = '#1C1C28';
                heading = '#E6E1C8';
                body = '#D8D3BA';
                break;
            case 'morandi':
                background = '#EAE3DC';
                heading = '#8A7968';
                body = '#7D7461';
                break;
            default:
                background = '#F9F9F9';
                heading = '#222222';
                body = '#444444';
        }
        
        document.getElementById('backgroundColor').value = background;
        document.getElementById('headingColor').value = heading;
        document.getElementById('bodyColor').value = body;
        
        // 如果已经生成了预览，则更新预览
        if (pages.length > 0) {
            updatePreview();
        }
    }
    
    // 获取字体系列值
    function getFontFamilyValue(fontName) {
        switch (fontName) {
            case '黑体':
                return '"黑体", "SimHei", sans-serif';
            case '宋体':
                return '"宋体", "SimSun", serif';
            case '楷体':
                return '"楷体", "KaiTi", cursive';
            case '微软雅黑':
                return '"微软雅黑", "Microsoft YaHei", sans-serif';
            case '苹方':
                return '"苹方", "PingFang SC", sans-serif';
            case '思源黑体':
                return '"思源黑体", "Source Han Sans", "Noto Sans CJK", sans-serif';
            case '思源宋体':
                return '"思源宋体", "Source Han Serif", "Noto Serif CJK", serif';
            case '华文黑体':
                return '"华文黑体", "STHeiti", sans-serif';
            case '华文宋体':
                return '"华文宋体", "STSong", serif';
            case '华文楷体':
                return '"华文楷体", "STKaiti", cursive';
            case '方正黑体':
                return '"方正黑体", "FZHei-B01S", sans-serif';
            case '方正宋体':
                return '"方正宋体", "FZShuSong-Z01S", serif';
            case 'Arial':
                return 'Arial, sans-serif';
            case 'Helvetica':
                return 'Helvetica, Arial, sans-serif';
            case 'Times New Roman':
                return '"Times New Roman", Times, serif';
            case 'Georgia':
                return 'Georgia, serif';
            case 'Courier New':
                return '"Courier New", Courier, monospace';
            default:
                return '"黑体", "SimHei", sans-serif';
        }
    }
    
    // 监听窗口大小改变
    window.addEventListener('resize', function() {
        if (pages.length > 0) {
            setTimeout(autoScalePreview, 100);
        }
    });
});