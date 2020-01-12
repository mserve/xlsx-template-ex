import * as moment from 'moment';
//import { CellRange } from './cell-range';
import { TemplateExpression, TemplatePipe } from './template-expression';
import { WorkSheetHelper } from './worksheet-helper';

// Excel Cell?

export class TemplateEngine {
  private readonly regExpBlocks: RegExp = /\[\[.+?]]/g;
  private readonly regExpValues: RegExp = /{{.+?}}/g;

  constructor(private wsh: WorkSheetHelper, private data: any) {
    console.log('TemplateEngine loaded with sheet ' + wsh.sheetName);
  }
  
  public async execute() {
    console.log('Starting TemplateEngine with sheet ' + this.wsh.sheetName);
    let usedRange = this.wsh.getUsedRange();
    await this.wsh.syncContext();
    this.processBlocks(usedRange, this.data);
    usedRange = this.wsh.getUsedRange();
    await this.wsh.syncContext();
    this.processValues(usedRange, this.data);
    console.log('TemplateEngine done for sheet ' + this.wsh.sheetName);
  }

  private processBlocks(cellRange: Excel.Range, data: any): Excel.Range {
    console.debug('Starting processBlocks');
    /* As this is a "real" range, it is always valid
    if (!cellRange.valid) {
      console.log(
        'xlsx-template-officejs: Process blocks failed.',
        'The cell range is invalid and will be skipped:',
        this.wsh.name, cellRange
      );
      return cellRange;
    }
    */
    let restart;
    do {
      restart = false;
      this.wsh.eachCell(cellRange, (cell: Excel.Range) => {        
        console.debug('eachCell on one cell')        
        let cVal = cell.values[0][0];
        if (typeof cVal !== "string") {
          return null;
        }
        const matches = (cVal as string).match(this.regExpBlocks);
        if (!Array.isArray(matches) || !matches.length) {
          return null;
        }

        matches.forEach((rawExpression: string) => {
          const tplExp = new TemplateExpression(rawExpression, rawExpression.slice(2, -2));
          cVal = (cVal as string).replace(tplExp.rawExpression, '');
          cell.values = [[cVal]];

          let resultData = data[tplExp.valueName];
          if (!data[tplExp.valueName] && this.data[tplExp.valueName]) {
            resultData = this.data[tplExp.valueName];
          }

          cellRange = this.processBlockPipes(cellRange, cell, tplExp.pipes, resultData);
        });

        restart = true;
        return false;
      });
    } while (restart);
    return cellRange;
  }

  private processValues(cellRange: Excel.Range, data: any): void {
    console.debug('Starting processValues');
    /*
    if (!cellRange.valid) {
      console.log(
        'xlsx-template-officejs: Process values failed.',
        'The cell range is invalid and will be skipped:',
        this.wsh.sheetName, cellRange
      );
      return;
    }
    */

    this.wsh.eachCell(cellRange, (cell: Excel.Range) => {
      let cVal = cell.values[0][0];
      if (typeof cVal !== "string") {
        return;
      }

      const matches = cVal.match(this.regExpValues);
      if (!Array.isArray(matches) || !matches.length) {
        return;
      }

      matches.forEach((rawExpression: string) => {
        const tplExp = new TemplateExpression(rawExpression, rawExpression.slice(2, -2));
        let resultValue: any = data[tplExp.valueName] || '';
        if (!data[tplExp.valueName] && this.data[tplExp.valueName]) {
          resultValue = this.data[tplExp.valueName];
        }
        resultValue = this.processValuePipes(cell, tplExp.pipes, resultValue);
        cVal = resultValue;
      });
      cell.values[0][0] = cVal;
    });
  }

  private processValuePipes(_cell: Excel.Range, pipes: TemplatePipe[], value: any): string {
    console.debug('Starting processValuePipes');
    try {
      pipes.forEach((pipe: TemplatePipe) => {
        switch (pipe.pipeName) {
          case 'date':
            // value = this.valuePipeDate(value, ...pipe.pipeParameters);
            value = this.valuePipeDate(value);
            break;
          /*case 'image':
            // value = this.valuePipeImage(cell, value, ...pipe.pipeParameters);
            value = this.valuePipeImage(cell, value);
            break;
            */
          case 'find':
            value = this.valuePipeFind(value, ...pipe.pipeParameters);
            break;
          case 'get':
            value = this.valuePipeGet(value, ...pipe.pipeParameters);
            break;
          case 'time':
            value = this.valuePipeTime(value);
            break;
          case 'datetime':
            value = this.valuePipeDateTime(value);
            break;
          case 'number':
            value = this.valuePipeNumber(value);
            break;
          default:
            value = 'xlsx-template-officejs: value of pipe not found:' + pipe.pipeName;
            console.warn(value);
        }
      });
    } catch (error) {
      console.error('xlsx-template-officejs: Error on process values of pipes', error);
      return 'xlsx-template-officejs: Error on process values of pipes. Look for more details in a console.';
    }

    return value || '';
  }

  private processBlockPipes(cellRange: Excel.Range, cell: Excel.Range, pipes: TemplatePipe[], data: any): Excel.Range {
    // console.log('bp', pipes, data);
    //ms const newRange = CellRange.createFromRange(cellRange);
    let insertedRows;
    try {
      pipes.forEach((pipe: TemplatePipe) => {
        switch (pipe.pipeName) {
          case 'repeat-rows':
            // insertedRows = this.blockPipeRepeatRows.apply(this, [cell, data].concat(pipe.pipeParameters));
            insertedRows += this.blockPipeRepeatRows(cell, data, ...pipe.pipeParameters);

            break;
          case 'tile':
            insertedRows += this.blockPipeTile(cell, data, ...pipe.pipeParameters);
            break;
          case 'filter':
            data = this.blockPipeFilter(data, ...pipe.pipeParameters);
            break;
          default:
            console.warn('xlsx-template-officejs: Block pipe not found:', pipe.pipeName, pipe.pipeParameters);
        }
      });
    } catch (error) {
      console.error('xlsx-template-officejs: Error on process a block of pipes', error);
      cell.values[0][0] = 'xlsx-template-officejs: Error on process a block of pipes. Look for more details in a console.';
    }
    return cellRange.getResizedRange(insertedRows, 0);
  }

  private valuePipeNumber(value?: any): any {
    if (Number(value) && value % 1 !== 0) {
      return parseFloat(value);
    } else if (Number(value) && value % 1 === 0) {
      return parseInt(value);
    }
    return value;
  }

  private valuePipeDate(date?: number | string): string {
    return date ? moment(new Date(date)).format('DD.MM.YYYY') : '';
  }


  private valuePipeTime(date?: number | string): string {
    return date ? moment(new Date(date)).format('HH:mm:ss') : '';
  }

  private valuePipeDateTime(date?: number | string): string {
    return date ? moment(new Date(date)).format('DD.MM.YYYY HH:mm:ss') : '';
  }

  /*
  private valuePipeImage(cell: Cell, fileName: string): string {
    if (fs.existsSync(fileName)) {
      this.wsh.addImage(fileName, cell);
      return fileName;
    }
    return ``;
  }
  */

  /** Find object in array by value of a property */
  private valuePipeFind(arrayData: any[], propertyName?: string, propertyValue?: string): any | null {
    if (Array.isArray(arrayData) && propertyName && propertyName) {
      return arrayData.find(item => item && item[propertyName] && item[propertyName].length > 0 && item[propertyName] == propertyValue);
    }
    return null;
  }

  private valuePipeGet(data: any[], propertyName?: string): any | null {
    return data && propertyName && data[propertyName] || null;
  }

  private blockPipeFilter(dataArray: any[], propertyName?: string, propertyValue?: string): any[] {
    if (Array.isArray(dataArray) && propertyName) {
      if (propertyValue) {
        return dataArray.filter(item => typeof item === "object" &&
          item[propertyName] &&
          item[propertyName].length > 0 &&
          item[propertyName] === propertyValue);
      }
      return dataArray.filter(item => typeof item === "object" &&
        item.hasOwnProperty(propertyName) &&
        item[propertyName] &&
        item[propertyName].length > 0
      );
    }
    return dataArray;
  }

  /** @return {number} count of inserted rows */
  blockPipeRepeatRows(cell: Excel.Range, dataArray: any[], countRows?: number | string): number {
    if (!Array.isArray(dataArray) || !dataArray.length) {
      console.warn('TemplateEngine.blockPipeRepeatRows', cell.address,
        'The data must be not empty array, but got:', dataArray
      );
      return 0;
    }
    let countRowsNum = +countRows > 0 ? +countRows : 1;
    const startRow = +cell.rowIndex;
    const endRow = startRow + countRowsNum - 1;
    if (dataArray.length > 1) {
      this.wsh.cloneRows(startRow, endRow, dataArray.length - 1);
    }

    const wsDimension = this.wsh.getUsedRange();


    // Range contains all data in those rose
    let sectionRange = this.wsh.getRange(startRow, wsDimension.columnIndex, endRow, wsDimension.columnIndex + wsDimension.columnCount);

    dataArray.forEach(data => {
      sectionRange = this.processBlocks(sectionRange, data);
      this.processValues(sectionRange, data);
      // Move range down
      sectionRange = sectionRange.getOffsetRange(countRowsNum, 0);
    });
    return (dataArray.length - 1) * countRowsNum;
  }

  /** @return {number} count of inserted rows */
  private blockPipeTile(cell: Excel.Range, dataArray: any[], blockRows?: number | string, blockColumns?: number | string,
    tileColumns?: number | string): number {
    // return;
    if (!Array.isArray(dataArray) || !dataArray.length) {
      console.warn('TemplateEngine.blockPipeTile', cell.address,
        'The data must be not empty array, but got:', dataArray
      );
      return 0;
    }

    blockRows = +blockRows > 0 ? +blockRows : 1;
    blockColumns = +blockColumns > 0 ? +blockColumns : 1;
    tileColumns = +tileColumns > 0 ? +tileColumns : 1;

    let blockRange = this.wsh.getRange(
      +cell.rowIndex, +cell.columnIndex, +cell.rowIndex + blockRows - 1, +cell.columnIndex + blockColumns - 1
    );

    const cloneRowsCount = Math.ceil(dataArray.length / tileColumns) - 1;
    if (dataArray.length > tileColumns) {
      this.wsh.cloneRows(blockRange.rowIndex, blockRange.rowIndex+blockRange.rowCount, cloneRowsCount);
    }

    let tileColumn = 1, tileRange = blockRange.getOffsetRange(0,0);
    dataArray.forEach((data, idx: number, array: any[]) => {
      // Prepare the next tile
      if ((idx !== array.length - 1) && (tileColumn + 1 <= tileColumns)) {
        const nextTileRange = tileRange.getOffsetRange(0, tileRange.columnCount);
        this.wsh.copyCellRange(tileRange, nextTileRange);
      }

      // Process templates
      tileRange = this.processBlocks(tileRange, data);
      this.processValues(tileRange, data);
      // Move tiles
      if (idx !== array.length - 1) {
        tileColumn++;
        if (tileColumn <= tileColumns) {
          tileRange = tileRange.getOffsetRange(0, tileRange.columnIndex);
        } else {
          tileColumn = 1;
          blockRange = tileRange.getOffsetRange(tileRange.columnCount, 0);
          tileRange = blockRange.getOffsetRange(0 ,0);
        }
      }
    });

    return cloneRowsCount * blockRange.rowCount;
  }
}
