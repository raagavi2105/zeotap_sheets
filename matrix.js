// Matrix and Data Management
class SpreadsheetMatrix {
    constructor(rows = 100, cols = 26) {
        this.rows = rows;
        this.cols = cols;
        this.matrix = this.createMatrix();
        this.undoStack = [];
        this.redoStack = [];
    }

    createMatrix() {
        const matrix = new Array(this.rows);
        for (let i = 0; i < this.rows; i++) {
            matrix[i] = new Array(this.cols).fill().map(() => ({
                value: '',
                formula: '',
                style: {},
                isProtected: false
            }));
        }
        return matrix;
    }

    getCellData(row, col) {
        return this.matrix[row][col];
    }

    setCellData(row, col, data) {
        this.undoStack.push({
            row,
            col,
            previousData: { ...this.matrix[row][col] }
        });
        this.matrix[row][col] = { ...this.matrix[row][col], ...data };
        this.redoStack = []; // Clear redo stack on new change
    }

    undo() {
        if (this.undoStack.length === 0) return;
        
        const lastChange = this.undoStack.pop();
        this.redoStack.push({
            row: lastChange.row,
            col: lastChange.col,
            previousData: { ...this.matrix[lastChange.row][lastChange.col] }
        });
        this.matrix[lastChange.row][lastChange.col] = lastChange.previousData;
    }

    redo() {
        if (this.redoStack.length === 0) return;
        
        const lastUndo = this.redoStack.pop();
        this.undoStack.push({
            row: lastUndo.row,
            col: lastUndo.col,
            previousData: { ...this.matrix[lastUndo.row][lastUndo.col] }
        });
        this.matrix[lastUndo.row][lastUndo.col] = lastUndo.previousData;
    }

    exportData() {
        return JSON.stringify({
            matrix: this.matrix,
            rows: this.rows,
            cols: this.cols
        });
    }

    importData(jsonData) {
        const data = JSON.parse(jsonData);
        this.rows = data.rows;
        this.cols = data.cols;
        this.matrix = data.matrix;
        this.undoStack = [];
        this.redoStack = [];
    }

    protected(row, col) {
        return this.matrix[row][col].isProtected;
    }

    toggleProtection(row, col) {
        this.setCellData(row, col, {
            isProtected: !this.matrix[row][col].isProtected
        });
    }
}

// Initialize the matrix
const spreadsheetMatrix = new SpreadsheetMatrix();