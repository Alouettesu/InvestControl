class ClosedOperationsSelector {
    constructor (sourceSheet, destSheet) {
        this.sourceSheet = sourceSheet;
        this.destSheet = destSheet;
        this.Columns = Object.freeze({"operationType":1, 
                                        "date":2, 
                                        "instrumentType":3,
                                        "isin": 4,
                                        "instrumentName": 5,
                                        "quantity": 6,
                                        "price": 7,
                                        "payment": 8,
                                        "currency": 9,
                                        "commission": 10
                                    });
    }

    rowToOperation(range) {
        var values = range.getValues();
        return {
            operationType: values[0][this.Columns.operationType - 1],
            date: values[0][this.Columns.date - 1],
            instrumentType: values[0][this.Columns.instrumentType - 1],
            isin: values[0][this.Columns.isin - 1],
            instrumentName: values[0][this.Columns.instrumentName - 1],
            quantity: values[0][this.Columns.quantity - 1],
            price: values[0][this.Columns.price - 1],
            payment: values[0][this.Columns.payment - 1],
            currency: values[0][this.Columns.currency - 1],
            commission: values[0][this.Columns.commission - 1]
        }
    }

    select() {
        var operationsInstruments = [];
        var firstRow = 2;
        var lastRow = this.sourceSheet.getDataRange().getLastRow();
        for (var row = firstRow; row < lastRow; row++) {
            var operation = this.rowToOperation(this.sourceSheet.getRange(row, 1, 1, Object.keys(this.Columns).length));
            if (!operationsInstruments.find(x => x.isin === operation.isin)) {
                operationsInstruments.push({isin: operation.isin, operations: Array(0)});
            }
            operationsInstruments.find(x => x.isin === operation.isin).operations.push(operation);
        }

        operationsInstruments.forEach(operations => {
            var sum = 0;
            var operationPrintIndex = 0;
            for (var i = 0; i < operations.operations.length; i++) {
                sum += operations.operations[i].quantity * (operations.operations[i].operationType === "Sell" ? -1 : 1);
                if (sum == 0)
                {
                    this.printOperations(operations.operations, operationPrintIndex, i);
                    operationPrintIndex = i + 1;
                }
            }
        });
    }

    printOperations(operationsArray, firstIndex, lastIndex) {
        var destRow = this.destSheet.getDataRange().getLastRow() + 1;
        for (var i = firstIndex; i <= lastIndex; i++) {
            Object.keys(operationsArray[i]).forEach(key => {
                var range = this.destSheet.getRange(destRow, this.Columns[key], 1, 1);
                range.setValue(operationsArray[i][key]);
            });
            destRow++;            
        }

    }
}