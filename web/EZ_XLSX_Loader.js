import { app } from "../../../scripts/app.js";

// Code largely inspired by FILL NODES, credit to the author: https://github.com/filliptm/ComfyUI_Fill-Nodes
// Adapted from EZ_CSV_Loader.js to support XLSX (Excel) files, including sheet selection.

app.registerExtension({
    name: "Comfy.EZ_XLSX_Loader",
    async nodeCreated(node) {
        if (node.comfyClass === "EZ_XLSX_Loader") {
            addXLSXBrowserUI(node);
        }
    }
});

async function addXLSXBrowserUI(node) {
    // Tweakable variables
    const CLICK_Y_OFFSET = 0;
    const CLICK_X_OFFSET = -2;

    const xlsxFileWidget = node.widgets.find(w => w.name === "xlsx_file");
    const selectedRowWidget = node.widgets.find(w => w.name === "selected_row");
    const selectionModeWidget = node.widgets.find(w => w.name === "selection_mode");
    const filterTextWidget = node.widgets.find(w => w.name === "filter_text");
    const sheetNameWidget = node.widgets.find(w => w.name === "sheet_name");

    if (!xlsxFileWidget || !selectedRowWidget || !selectionModeWidget || !filterTextWidget) {
        console.error("Required widgets not found:", { xlsxFileWidget, selectedRowWidget, selectionModeWidget, filterTextWidget });
        return;
    }

    xlsxFileWidget.hidden = false;
    selectedRowWidget.hidden = true;
    selectionModeWidget.hidden = false;
    filterTextWidget.hidden = false;
    if (sheetNameWidget) {
        sheetNameWidget.hidden = false;
    }

    const MIN_WIDTH = 310;
    const MIN_HEIGHT = 400;
    const TOP_PADDING = 250;
    const BOTTOM_PADDING = 5;
    const BOTTOM_SKIP = 10;
    const TOP_BAR_HEIGHT = 0;
    const ROW_HEIGHT = 28;
    const ROW_PADDING = 5;
    const EXTRA_ROW_PADDING = 2;
    const SCROLLBAR_WIDTH = 13;
    const MIN_COLUMN_WIDTH = 150; // Minimum width for a column
    const TEXT_PADDING = 10; // Padding for text within row
    const PREVIEW_PADDING = 20; // Padding for preview text
    const PREVIEW_SKIP = 152; // Skip for preview text
    const HEADERS_SKIP = 225; // Skip for headers preview
    const BORDER_RADIUS = 0;
    const SELECTION_BORDER_RADIUS = 0;
    const SELECTION_BORDER_PADDING = 0;
    const ELLIPSIS = "...";
    const HEADERS_LINE_HEIGHT = 14; // Adjustable line height for headers preview

    const COLORS = {
        background: "#1e1e1e",
        topBar: "#252526",
        row: "#2d2d30",
        rowHover: "#3e3e42",
        rowSelected: "#0e639c",
        text: "#ffffff",
        headers: "#a9a9a9",
        scrollbar: "#3e3e42",
        scrollbarHover: "#505050",
        divider: "#4f0074",
        dividerHover: "#16727c"
    };

    let currentFile = null;
    let filterText = filterTextWidget.value;
    let selectedRows = new Set();
    let headers = [];
    let rows = [];
    let scrollOffset = 0;
    let isDragging = false;
    let scrollStartY = 0;
    let scrollStartOffset = 0;

    async function updateRows() {
        try {
            const response = await fetch('/ez_xlsx_browser/get_directory_structure', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ path: currentFile, filter: filterText, sheet: sheetNameWidget ? sheetNameWidget.value : "" })
            });

            if (!response.ok) {
                const errorData = await response.json();
                console.error("Server error:", errorData.error);
                return;
            }

            const data = await response.json();
            if (!data.rows) {
                console.error("Invalid response format:", data);
                return;
            }

            headers = data.headers || [];
            rows = data.rows || [];
            node.setDirtyCanvas(true);
        } catch (error) {
            console.error("Error updating rows:", error);
        }
    }

    async function fetchFileInfo(relativePath) {
        try {
            const response = await fetch('/ez_xlsx_browser/get_file_info', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ relative_path: relativePath })
            });

            if (!response.ok) {
                const errorData = await response.json();
                console.error("Server error:", errorData.error);
                return null;
            }

            const result = await response.json();
            return result.full_path || null;
        } catch (error) {
            console.error("Error fetching file info:", error);
            return null;
        }
    }

    function updateSelectedRows(rowIdx) {
        if (selectionModeWidget.value === "multiple" || selectionModeWidget.value === "random") {
            if (selectedRows.has(rowIdx)) {
                selectedRows.delete(rowIdx);
            } else {
                selectedRows.add(rowIdx);
            }
        } else {
            selectedRows.clear();
            selectedRows.add(rowIdx);
        }
        const selectedRowsString = Array.from(selectedRows).join(",");
        selectedRowWidget.value = selectedRowsString;
        node.setDirtyCanvas(true);
    }

    function drawHeadersPreview(ctx) {
        ctx.fillStyle = COLORS.headers;
        ctx.font = "11px Arial";
        const maxWidth = node.size[0] - PREVIEW_PADDING * 2 - 2;

        // First line: Show "N headers:" where N is the number of headers
        const headerCount = headers.length;
        ctx.fillText(`${headerCount} Headers:`, PREVIEW_PADDING, HEADERS_SKIP);

        // Second line: Headers separated by " | "
        if (headers.length > 0) {
            let headerString = headers.join(" | ");

            // If too long, truncate with ellipsis
            if (ctx.measureText(headerString).width > maxWidth) {
                let truncated = headerString;
                while (ctx.measureText(truncated + ELLIPSIS).width > maxWidth && truncated.length > 0) {
                    truncated = truncated.slice(0, -1);
                }
                headerString = truncated + ELLIPSIS;
            }

            ctx.fillText(headerString, PREVIEW_PADDING, HEADERS_SKIP + HEADERS_LINE_HEIGHT);
        }
    }

    function drawPreviewText(ctx) {
        ctx.fillStyle = COLORS.text;
        ctx.font = "12px Arial";
        const maxWidth = node.size[0] - PREVIEW_PADDING * 2;
        let displayText = "";
        if (selectionModeWidget.value === "random") {
            const rowCount = selectedRows.size > 0 ? selectedRows.size : rows.length;
            displayText = `selecting from ${rowCount} rows`;
        } else if (selectionModeWidget.value === "multiple") {
            displayText = `${selectedRows.size} rows selected`;
        } else if (selectionModeWidget.value === "single" && selectedRows.size === 1 && headers.length > 0) {
            const idx = Array.from(selectedRows)[0];
            if (rows[idx]) {
                displayText = rows[idx][0] || `Row ${idx+1}`;
            }
        }
        if (ctx.measureText(displayText).width > maxWidth) {
            let truncatedText = displayText;
            while (ctx.measureText(truncatedText + ELLIPSIS).width > maxWidth && truncatedText.length > 0) {
                truncatedText = truncatedText.slice(0, -1);
            }
            displayText = truncatedText + ELLIPSIS;
        }
        ctx.fillText(displayText, PREVIEW_PADDING, PREVIEW_SKIP);
    }

    const refreshButton = node.addWidget("button", "Refresh / Clear", null, () => {
        selectedRows.clear();
        selectedRowWidget.value = "";
        (async () => {
            currentFile = await fetchFileInfo(xlsxFileWidget.value);
            if (currentFile) {
                updateRows();
            }
        })();
    });

    xlsxFileWidget.callback = () => {
        selectedRows.clear();
        selectedRowWidget.value = "";
        (async () => {
            currentFile = await fetchFileInfo(xlsxFileWidget.value);
            if (currentFile) {
                updateRows();
            }
        })();
    };

    filterTextWidget.callback = () => {
        filterText = filterTextWidget.value;
        updateRows();
    };

    if (sheetNameWidget) {
        sheetNameWidget.callback = () => {
            selectedRows.clear();
            selectedRowWidget.value = "";
            updateRows();
        };
    }

    // Add Invert Selection button widget
    const invertButton = node.addWidget("button", "Invert Selection", null, () => {
        if (selectionModeWidget.value !== "multiple" && selectionModeWidget.value !== "random") return;
        const allRowIndices = Array.from({length: rows.length}, (_, i) => i);
        const newSelected = new Set();
        for (const idx of allRowIndices) {
            if (!selectedRows.has(idx)) newSelected.add(idx);
        }
        selectedRows = newSelected;
        // Update widget value
        const selectedRowsString = Array.from(selectedRows).join(",");
        selectedRowWidget.value = selectedRowsString;
        node.setDirtyCanvas(true);
    });
    // Ensure correct enabled state on load
    setTimeout(() => {
        invertButton.disabled = selectionModeWidget.value !== "multiple" && selectionModeWidget.value !== "random";
    }, 0);

    // Update button enabled/disabled state on mode change
    const origSelectionModeCallback = selectionModeWidget.callback;
    selectionModeWidget.callback = () => {
        selectedRows.clear();
        selectedRowWidget.value = "";
        invertButton.disabled = selectionModeWidget.value !== "multiple" && selectionModeWidget.value !== "random";
        node.setDirtyCanvas(true);
        if (origSelectionModeCallback) origSelectionModeCallback();
    };

    node.onDrawBackground = function(ctx) {
        if (!this.flags.collapsed) {
            const pos = TOP_PADDING - TOP_BAR_HEIGHT;
            ctx.fillStyle = COLORS.background;
            ctx.fillRect(0, pos, this.size[0], this.size[1] - pos - BOTTOM_SKIP);
            ctx.fillStyle = COLORS.topBar;
            ctx.fillRect(0, pos, this.size[0], TOP_BAR_HEIGHT);
            drawHeadersPreview(ctx);
            drawPreviewText(ctx);
            ctx.save();
            ctx.beginPath();
            ctx.rect(0, TOP_PADDING, this.size[0] - SCROLLBAR_WIDTH, this.size[1] - TOP_PADDING - BOTTOM_PADDING - BOTTOM_SKIP);
            ctx.clip();
            drawRows(ctx, 0, TOP_PADDING - scrollOffset, this.size[0] - SCROLLBAR_WIDTH - 10, this.size[1] - TOP_PADDING - BOTTOM_PADDING - BOTTOM_SKIP);
            ctx.restore();
            drawScrollbar(ctx, this.size[0] - SCROLLBAR_WIDTH, TOP_PADDING, SCROLLBAR_WIDTH, this.size[1] - TOP_PADDING - BOTTOM_PADDING - BOTTOM_SKIP, scrollOffset, getTotalRowsHeight());
        }
    };

    function drawRoundedRect(ctx, x, y, width, height, radius, color) {
        ctx.beginPath();
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + width - radius, y);
        ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
        ctx.lineTo(x + width, y + height - radius);
        ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        ctx.lineTo(x + radius, y + height);
        ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
        ctx.lineTo(x, y + radius);
        ctx.quadraticCurveTo(x, y, x + radius, y);
        ctx.closePath();
        ctx.fillStyle = color;
        ctx.fill();
    }

    function drawScrollbar(ctx, x, y, width, height, offset, totalHeight) {
        drawRoundedRect(ctx, x, y, width, height, width / 2, COLORS.scrollbar);
        const visibleHeight = height;
        const scrollHeight = Math.max(height * (visibleHeight / totalHeight), 20);
        const maxOffset = Math.max(0, totalHeight - visibleHeight);
        const scrollY = y + (offset / maxOffset) * (height - scrollHeight);
        drawRoundedRect(ctx, x, scrollY, width, scrollHeight, width / 2, COLORS.scrollbarHover);
    }

    function getTotalRowsHeight() {
        const columns = Math.max(2, Math.floor((node.size[0] - SCROLLBAR_WIDTH) / MIN_COLUMN_WIDTH));
        const rowCount = rows.length;
        const rowsPerCol = Math.ceil(rowCount / columns);
        return rowsPerCol * (ROW_HEIGHT + ROW_PADDING);
    }

    function drawRows(ctx, x, y, width, height) {
        ctx.fillStyle = COLORS.background;
        ctx.fillRect(x, y, width, height);
        const columns = Math.max(2, Math.floor((node.size[0] - SCROLLBAR_WIDTH) / MIN_COLUMN_WIDTH));
        const columnWidth = (width - ROW_PADDING * (columns + 1)) / columns;
        const rowCount = rows.length;
        const rowsPerCol = Math.ceil(rowCount / columns);
        const visibleHeight = height;
        const startRow = Math.floor(scrollOffset / (ROW_HEIGHT + ROW_PADDING));
        const endRow = Math.min(rowsPerCol, startRow + Math.ceil(visibleHeight / (ROW_HEIGHT + ROW_PADDING)) + 2);
        for (let row = startRow; row < endRow; row++) {
            for (let col = 0; col < columns; col++) {
                const rowIndex = row * columns + col;
                if (rowIndex >= rows.length) break;
                const rowData = rows[rowIndex];
                const label = rowData[0] || `Row ${rowIndex+1}`;
                const xPos = x + EXTRA_ROW_PADDING + ROW_PADDING + col * (columnWidth + ROW_PADDING);
                const yPos = y + EXTRA_ROW_PADDING + row * (ROW_HEIGHT + ROW_PADDING) + ROW_PADDING;
                // Draw row background
                const bgColor = selectedRows.has(rowIndex) ? COLORS.rowSelected : COLORS.row;
                drawRoundedRect(ctx, xPos, yPos, columnWidth, ROW_HEIGHT, BORDER_RADIUS, bgColor);
                // Draw row label with truncation
                ctx.fillStyle = COLORS.text;
                ctx.font = "12px Arial";
                const maxTextWidth = columnWidth - TEXT_PADDING * 2;
                const textMetrics = ctx.measureText(label);
                let displayText = label;
                if (textMetrics.width > maxTextWidth) {
                    let truncatedText = label;
                    while (ctx.measureText(truncatedText + ELLIPSIS).width > maxTextWidth && truncatedText.length > 0) {
                        truncatedText = truncatedText.slice(0, -1);
                    }
                    displayText = truncatedText + ELLIPSIS;
                }
                ctx.fillText(displayText, xPos + TEXT_PADDING, yPos + ROW_HEIGHT / 2 + 4);
            }
        }
    }

    node.onMouseDown = function(event) {
        const pos = TOP_PADDING - TOP_BAR_HEIGHT;
        const localY = event.canvasY - this.pos[1] - pos + CLICK_Y_OFFSET;
        const localX = event.canvasX - this.pos[0] + CLICK_X_OFFSET;
        if (localY < 0 || localY > this.size[1] || localX < 0 || localX > this.size[0]) {
            return false;
        }
        if (localY > TOP_BAR_HEIGHT && localY < this.size[1] - pos - 10) {
            if (localX >= 0 && localX < this.size[0] - SCROLLBAR_WIDTH) {
                // Calculate which row was clicked
                const columns = Math.max(2, Math.floor((this.size[0] - SCROLLBAR_WIDTH) / MIN_COLUMN_WIDTH));
                const column = Math.floor(localX / ((this.size[0] - SCROLLBAR_WIDTH) / columns));
                const rowsPerCol = Math.ceil(rows.length / columns);
                const row = Math.floor((localY - TOP_BAR_HEIGHT + scrollOffset) / (ROW_HEIGHT + ROW_PADDING));
                const rowIndex = row * columns + column;
                if (rowIndex >= 0 && rowIndex < rows.length) {
                    updateSelectedRows(rowIndex);
                }
                return true;
            } else if (localX >= this.size[0] - SCROLLBAR_WIDTH) {
                // Click on scrollbar
                isDragging = true;
                scrollStartY = event.canvasY;
                scrollStartOffset = scrollOffset;
                return true;
            }
        }
        return false;
    };

    node.onMouseMove = function(event) {
        const pos = TOP_PADDING - TOP_BAR_HEIGHT;
        const localY = event.canvasY - this.pos[1] - pos + CLICK_Y_OFFSET;
        const localX = event.canvasX - this.pos[0] + CLICK_X_OFFSET;
        if (isDragging) {
            const totalHeight = getTotalRowsHeight();
            const visibleHeight = this.size[1] - TOP_PADDING - BOTTOM_PADDING - BOTTOM_SKIP;
            const maxOffset = Math.max(0, totalHeight - visibleHeight);
            const scrollMove = (event.canvasY - scrollStartY) * (totalHeight / visibleHeight);
            scrollOffset = Math.max(0, Math.min(maxOffset, scrollStartOffset + scrollMove));
            this.setDirtyCanvas(true);
            return true;
        }
        return false;
    };

    node.onMouseUp = function(event) {
        isDragging = false;
        document.body.style.cursor = 'default';
        return false;
    };

    function updateNodeSize() {
        const width = Math.max(MIN_WIDTH, node.size[0]);
        const height = Math.max(MIN_HEIGHT, node.size[1]);
        node.size[0] = width;
        node.size[1] = height;
    }

    node.onResize = function() {
        updateNodeSize();
        this.setDirtyCanvas(true);
    };

    // Restore selection from widget value on load
    setTimeout(() => {
        selectedRows = new Set(
            selectedRowWidget.value.split(',').map(t => t.trim()).filter(t => t !== '').map(t => parseInt(t)).filter(Number.isFinite)
        );
        invertButton.disabled = selectionModeWidget.value !== "multiple" && selectionModeWidget.value !== "random";
        node.setDirtyCanvas(true);
    }, 0);

    // Initialize
    setTimeout(async () => {
        currentFile = await fetchFileInfo(xlsxFileWidget.value);
        if (currentFile) {
            updateRows();
        }
        updateNodeSize();
    }, 0);
}
