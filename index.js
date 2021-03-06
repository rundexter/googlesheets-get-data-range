var _ = require('lodash'),
    Spreadsheet = require('edit-google-spreadsheet');

module.exports = {
    checkAuthOptions: function (step, dexter) {

        if(!step.input('LastRow').first() || !step.input('LastColumn').first())
            return 'A [worksheet, LastRow, LastColumn] inputs variable is required for this module';

        if(!dexter.environment('google_spreadsheet'))
            return 'A [google_spreadsheet] environment variable is required for this module';

        return false;
    },

    convertColumnLetter: function(val) {
        var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
            result = 0;

        var i, j;

        for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {

            result += Math.pow(base.length, j) * (base.indexOf(val[i]) + 1);
        }

        return result;
    },

    pickSheetData: function (rows, startRow, startColumn, numRows, numColumns) {
        var sheetData = {};

        for (var rowKey = startRow; rowKey <= numRows; rowKey++) {

            for (var columnKey = startColumn; columnKey <= numColumns; columnKey++) {

                if (rows[rowKey] !== undefined && rows[rowKey][columnKey] !== undefined) {

                    sheetData[rowKey] = sheetData[rowKey] || {};
                    sheetData[rowKey][columnKey] = rows[rowKey][columnKey];
                }
            }
        }

        return sheetData;
    },

    /**
     * The main entry point for the Dexter module
     *
     * @param {AppStep} step Accessor for the configuration for the step using this module.  Use step.input('{key}') to retrieve input data.
     * @param {AppData} dexter Container for all data used in this workflow.
     */
    run: function(step, dexter) {
        var credentials = dexter.provider('google').credentials(),
            error = this.checkAuthOptions(step, dexter);
        var spreadsheetId = dexter.environment('google_spreadsheet'),
            lastRow = step.input('LastRow').first(),
            lastColumn = step.input('LastColumn').first();

        var worksheetId = step.input('worksheet', 1).first(),
            startRow = 1,
            startColumn = 1,
            numRows = _.parseInt(lastRow),
            numColumns = this.convertColumnLetter(lastColumn);

        if (error) return this.fail(error);

        Spreadsheet.load({
            spreadsheetId: spreadsheetId,
            worksheetId: worksheetId,
            accessToken: {
                type: 'Bearer',
                token: _.get(credentials, 'access_token')
            }
        }, function (err, spreadsheet) {

            if (err)
                this.fail(err);
            else
                spreadsheet.receive({ getValues: true }, function (error, rows) {

                    error? this.fail(error) : this.complete({sheet: this.pickSheetData(rows, startRow, startColumn, numRows, numColumns)});
                }.bind(this));
        }.bind(this));

    }
};
