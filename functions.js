// Advanced Spreadsheet Functions
class SpreadsheetFunctions {
    static evaluateFormula(formula, matrix) {
        try {
            // Remove the = sign and evaluate the formula
            const expression = formula.substring(1);
            
            // Replace cell references with values
            const evaluatedExpression = expression.replace(/[A-Z]+[0-9]+/g, (match) => {
                const col = match.match(/[A-Z]+/)[0];
                const row = parseInt(match.match(/[0-9]+/)[0]) - 1;
                const colIndex = SpreadsheetFunctions.columnToIndex(col);
                return matrix.getCellData(row, colIndex).value || 0;
            });

            return eval(evaluatedExpression);
        } catch (error) {
            return '#ERROR!';
        }
    }

    static columnToIndex(column) {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + column.charCodeAt(i) - 64;
        }
        return index - 1;
    }

    static indexToColumn(index) {
        let column = '';
        while (index >= 0) {
            column = String.fromCharCode(65 + (index % 26)) + column;
            index = Math.floor(index / 26) - 1;
        }
        return column;
    }

    // Mathematical Functions
    static sum(range) {
        if (!range || range.length === 0) return '#ERROR!';
        return range.reduce((acc, cell) => acc + (parseFloat(cell.value) || 0), 0);
    }

    static average(range) {
        if (!range || range.length === 0) return '#ERROR!';
        return this.sum(range) / range.length;
    }

    static count(range) {
        if (!range) return '#ERROR!';
        return range.filter(cell => !isNaN(parseFloat(cell.value))).length;
    }

    static min(range) {
        if (!range || range.length === 0) return '#ERROR!';
        const validValues = range.map(cell => parseFloat(cell.value)).filter(val => !isNaN(val));
        return validValues.length ? Math.min(...validValues) : '#ERROR!';
    }

    static max(range) {
        if (!range || range.length === 0) return '#ERROR!';
        const validValues = range.map(cell => parseFloat(cell.value)).filter(val => !isNaN(val));
        return validValues.length ? Math.max(...validValues) : '#ERROR!';
    }

    // Statistical Functions
    static median(range) {
        if (!range || range.length === 0) return '#ERROR!';
        const values = range
            .map(cell => parseFloat(cell.value))
            .filter(val => !isNaN(val))
            .sort((a, b) => a - b);
        if (values.length === 0) return '#ERROR!';
        const mid = Math.floor(values.length / 2);
        return values.length % 2 ? values[mid] : (values[mid - 1] + values[mid]) / 2;
    }

    static variance(range) {
        if (!range || range.length === 0) return '#ERROR!';
        const avg = this.average(range);
        if (avg === '#ERROR!') return '#ERROR!';
        const squareDiffs = range.map(cell => {
            const diff = (parseFloat(cell.value) || 0) - avg;
            return diff * diff;
        });
        return this.sum(squareDiffs.map(diff => ({ value: diff }))) / range.length;
    }

    static standardDeviation(range) {
        const var_value = this.variance(range);
        return var_value === '#ERROR!' ? '#ERROR!' : Math.sqrt(var_value);
    }

    // Text Functions
    static concatenate(range) {
        if (!range) return '#ERROR!';
        return range.map(cell => cell.value || '').join('');
    }

    static left(text, n) {
        if (typeof text !== 'string' || !n) return '#ERROR!';
        return text.substring(0, n);
    }

    static right(text, n) {
        if (typeof text !== 'string' || !n) return '#ERROR!';
        return text.substring(text.length - n);
    }

    static mid(text, start, n) {
        if (typeof text !== 'string' || !start || !n) return '#ERROR!';
        return text.substring(start - 1, start - 1 + n);
    }

    static len(text) {
        if (typeof text !== 'string') return '#ERROR!';
        return text.length;
    }

    static proper(text) {
        if (typeof text !== 'string') return '#ERROR!';
        return text.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    }

    static trim(text) {
        if (typeof text !== 'string') return '#ERROR!';
        return text.trim();
    }

    // Data Quality Functions
    static removeDuplicates(range) {
        if (!range || range.length === 0) return '#ERROR!';
        const seen = new Set();
        return range.map(cell => {
            if (seen.has(cell.value)) {
                return { ...cell, value: '' };
            }
            seen.add(cell.value);
            return cell;
        });
    }

    static findReplace(range, find, replace) {
        if (!range || !find) return '#ERROR!';
        return range.map(cell => ({
            ...cell,
            value: (cell.value || '').replace(new RegExp(find, 'g'), replace || '')
        }));
    }

    // Logical Functions
    static if(condition, trueValue, falseValue) {
        try {
            return condition ? trueValue : falseValue;
        } catch (error) {
            return '#ERROR!';
        }
    }

    static and(...conditions) {
        try {
            return conditions.every(condition => condition);
        } catch (error) {
            return '#ERROR!';
        }
    }

    static or(...conditions) {
        try {
            return conditions.some(condition => condition);
        } catch (error) {
            return '#ERROR!';
        }
    }

    // Date Functions
    static now() {
        return new Date();
    }

    static today() {
        const date = new Date();
        return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }

    static dateFormat(date, format) {
        if (!(date instanceof Date)) return '#ERROR!';
        try {
            const options = {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit'
            };
            return date.toLocaleDateString(undefined, options);
        } catch (error) {
            return '#ERROR!';
        }
    }

    // Validation Functions
    static isNumeric(value) {
        return !isNaN(parseFloat(value)) && isFinite(value);
    }

    static isDate(value) {
        const date = new Date(value);
        return date instanceof Date && !isNaN(date);
    }

    static isEmail(value) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(value);
    }
}

// Export the class for use in other files
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpreadsheetFunctions;
}