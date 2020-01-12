import * as _ from 'lodash';

// import {Cell, Row, ValueType, Workbook, Worksheet} from "exceljs";
// import { CellRange } from "./cell-range";

/**
 * Callback for iterate cells
 * @return false - whether to break iteration
 */
export type iterateCells = (cell: Excel.Range) => void | false;

/**
 * Single Cell representation (i.e. range)
 * 
 */

export class WorkSheetHelper {

  constructor(private worksheet: Excel.Worksheet, private context: Excel.RequestContext) {
    worksheet.load('name');
    this.context.sync().then(() => {
      console.log('WorkSheetHelper ready for sheet ' + worksheet.name);
    });
    console.log('After then');
  }



  get sheetName() {
    return this.worksheet.name;
  }

  public async syncContext() {
    await this.context.sync();
  }

  /*
  public addImage(fileName: string, cell: Cell): void {
    const imgId = this.workbook.addImage({filename: fileName, extension: 'jpeg'});

    const cellRange = this.getMergeRange(cell);
    if (cellRange) {
      this.worksheet.addImage(imgId, {
        tl: {col: cellRange.left - 0.99999, row: cellRange.top - 0.99999},
        br: {col: cellRange.right, row: cellRange.bottom}
      });
    } else {
      this.worksheet.addImage(imgId, {
        tl: {col: +cell.col - 0.99999, row: +cell.row - 0.99999}, 
        br: {col: +cell.col, row: +cell.row},
      });
    }
  }
  */

  public getUsedRange() {
    console.debug('Starting getUsedRange');

    let usedRange = this.worksheet.getUsedRangeOrNullObject();
    if (!!usedRange) {
      usedRange.load(['columnCount', 'rowCount', 'rowIndex', 'columnIndex', 'values']);
      this.syncContext();
    }
    return usedRange;
  }

  public getRange(top: number, left: number, bottom: number, right: number) {
    console.debug('Starting getUsedRange');
    return this.worksheet.getRangeByIndexes(top, left, bottom - top, right - left);
  }
  /*
    public getSheetDimension(): CellRange {
      const dm = this.worksheet.getUsedRange()
      return new CellRange(dm.rowIndex, dm.columnIndex, dm.rowCount + dm.rowIndex, dm.columnCount + dm.columnIndex);
    }
  */
  public cloneRows(srcRowStart: number, srcRowEnd: number, countClones: number = 1): void {
    console.debug('Starting cloneRows');
    const countRows = srcRowEnd - srcRowStart + 1;
    const dxRow = countRows * countClones;
    const lastRow = this.getUsedRange().getLastRow().rowIndex + dxRow;

    // Move rows below   
    for (let rowSrcNumber = lastRow; rowSrcNumber > srcRowEnd; rowSrcNumber--) {
      const rowSrc = this.worksheet.getCell(rowSrcNumber, 1).getEntireRow();
      const rowDest = this.worksheet.getCell(rowSrcNumber + dxRow, 1).getEntireRow();
      this.moveRow(rowSrc, rowDest);
    }

    // Clone target rows
    for (let rowSrcNumber = srcRowEnd; rowSrcNumber >= srcRowStart; rowSrcNumber--) {
      const rowSrc = this.worksheet.getCell(rowSrcNumber, 1).getEntireRow();
      for (let cloneNumber = countClones; cloneNumber > 0; cloneNumber--) {
        const rowDest = this.worksheet.getCell(rowSrcNumber + countRows * cloneNumber, 1).getEntireRow();
        this.copyRow(rowSrc, rowDest);
      }
    }
  }

  /* Copy cell range to another range having same dimension */
  public copyCellRange(rangeSrc: Excel.Range, rangeDest: Excel.Range): void {
    console.debug('Starting copyCellRange');

    if (rangeSrc.rowCount !== rangeDest.rowCount || rangeSrc.columnCount !== rangeDest.columnCount) {
      console.warn('WorkSheetHelper.copyCellRange',
        'The cell ranges must have an equal size', rangeSrc, rangeDest
      );
      return;
    }
    rangeDest.values = rangeSrc.values;
    // todo: check intersection in the CellRange class
    /*
    const dRow = rangeDest.bottom - rangeSrc.bottom;
    const dCol = rangeDest.right - rangeSrc.right;
    this.eachCellReverse(rangeSrc, (cellSrc: Cell) => {
      const cellDest = this.worksheet.getCell(cellSrc.range.rowIndex + dRow, cellSrc.range.columnIndex + dCol);
      this.copyCell(cellSrc, cellDest);
    });
    */
  }


  /** Iterate cells from the left of the top to the right of the bottom */
  public eachCell(cellRange: Excel.Range, callBack: iterateCells) {
    console.debug('Starting eachCell');
    try {
      cellRange.load(['columnCount', 'rowCount']);
    } catch (_e) {
      console.error(' Loading silently failed');
      console.error(_e);
    }

    console.debug('Ready for promise');
    this.context.sync().then(() => {
      console.debug(`Range size: ${cellRange.rowCount} /  ${cellRange.columnCount} `);
      for (let r = 0; r < cellRange.rowCount; r++) {
        for (let c = 0; c < cellRange.columnCount; c++) {
          const cell = cellRange.getCell(r, c);
          if (cell) {
            if (callBack(cell) === false) {
              return;
            }
          }
        }
      }
    }).catch(() => {
      console.error('Failed to sync')
    });
    /*
    for (let r = cellRange.top; r <= cellRange.bottom; r++) {
      const row = this.worksheet.findRow(r);
      if (row) {
        for (let c = cellRange.left; c <= cellRange.right; c++) {
          const cell = row.findCell(c);
          if (cell && cell.type !== ValueType.Merge) {
            if (callBack(cell) === false) {
              return;
            }
          }
        }
      }
    }*/
  }

  /** Iterate cells from the right of the bottom to the top of the left */
  public eachCellReverse(cellRange: Excel.Range, callBack: iterateCells) {
    console.debug('Starting eachCellReverse');
    for (let r = cellRange.rowCount - 1; r >= 0; r--) {
      for (let c = cellRange.columnCount - 1; c >= 0; c--) {
        const cell = cellRange.getCell(r, c);
        if (cell) {
          if (callBack(cell) === false) {
            return;
          }
        }
      }
    }
  }

  /*
private getMergeRange(cell: Cell): CellRange {
  if (cell.isMerged && Array.isArray(this.worksheet.model['merges'])) {
    const address = cell.type === ValueType.Merge ? cell.master.address : cell.address;
    const cellRangeStr = this.worksheet.model['merges']
      .find((item: string) => item.indexOf(address + ':') !== -1);
    if (cellRangeStr) {
      const [cellTlAdr, cellBrAdr] = cellRangeStr.split(':', 2);
      return CellRange.createFromCells(
        this.worksheet.getCell(cellTlAdr),
        this.worksheet.getCell(cellBrAdr)
      );
    }
  }
  return null;
}
*/

  private moveRow(rowSrc: Excel.Range, rowDest: Excel.Range): void {
    this.copyRow(rowSrc, rowDest);
    this.clearRow(rowSrc);
  }

  private copyRow(rowSrc: Excel.Range, rowDest: Excel.Range): void {
    if (!(rowSrc.rowCount == 1 && rowDest.rowCount == 1))
      return;
    rowDest.getEntireRow().values = rowSrc.getEntireRow().values;
  }
  /*
    private copyCell(cellSrc: Excel.Range, cellDest: Excel.Range): void {
      if (!(cellSrc.rowCount == 1 && cellDest.rowCount == 1 && cellSrc.columnCount == 1 && cellDest.columnCount == 1))
        return;
      cellDest.values = cellSrc.values;
    }
  */
  private clearRow(row: Excel.Range): void {
    row.clear(Excel.ClearApplyTo.all);
  }

  // private clearCell(cell: Cell): void {
  //   cell.model = {
  //     address: cell.fullAddress.address, style: undefined, type: undefined, text: undefined, hyperlink: undefined,
  //     value: undefined, master: undefined, formula: undefined, sharedFormula: undefined, result: undefined
  //   };
  // }
}
