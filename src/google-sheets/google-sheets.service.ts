// src/google-sheets/google-sheets.service.ts
import {
  HttpException,
  HttpStatus,
  Injectable,
  OnModuleInit,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { HttpService } from '@nestjs/axios';
import { catchError, firstValueFrom } from 'rxjs';
import { AxiosError } from 'axios';
import { google, sheets_v4 } from 'googleapis';
import { JWT } from 'google-auth-library';
import * as path from 'path';

// DTO interface for incoming data
export interface InternalApiResponse {
  ObjectName: string;
  Year: number;
  Month: number;
  Plan: number;
  Fact: number;
}

@Injectable()
export class GoogleSheetsService implements OnModuleInit {
  private sheetsClient: sheets_v4.Sheets;

  constructor(
    private configService: ConfigService,
    private readonly httpService: HttpService,
  ) {}

  // Step 0: Initialize Google Sheets API client
  async onModuleInit() {
    const auth = new google.auth.GoogleAuth({
      keyFile: path.join(
        __dirname,
        '../../carbide-ratio-457005-s9-a191b5f03e01.json',
      ),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    this.sheetsClient = google.sheets({
      version: 'v4',
      auth: authClient as JWT,
    });
  }

  // Main method to fetch data and write to appropriate sheet
  async writeDataToGoogleSheets() {
    // Step 1.1: Load config variables
    const internalApiUrl = this.configService.get<string>('INTERNAL_API_URL');
    const spreadsheetId = this.configService.get<string>('SPREADSHEET_ID');

    // Step 1.2: Check configs
    if (!internalApiUrl) {
      throw new HttpException(
        'INTERNAL_API_URL is not defined in the config',
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
    if (!spreadsheetId) {
      throw new HttpException(
        'SPREADSHEET_ID is not defined in the config',
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }

    // Step 1.3: Fetch data from internal API
    const { data } = await firstValueFrom(
      this.httpService.get<InternalApiResponse[]>(internalApiUrl).pipe(
        catchError((error: AxiosError) => {
          throw new HttpException(
            `Failed to fetch from internal API: ${error.message}`,
            HttpStatus.BAD_GATEWAY,
          );
        }),
      ),
    );

    // Step 2: Group data by sheet name (ВА, Б)
    const groupedBySheet = new Map<string, InternalApiResponse[]>();
    for (const item of data) {
      const sheetName = item.ObjectName.startsWith('Б') ? 'Б' : 'ВА';
      if (!groupedBySheet.has(sheetName)) groupedBySheet.set(sheetName, []);
      groupedBySheet.get(sheetName)!.push(item);
    }
    console.log('Grouped data by sheet:', Array.from(groupedBySheet.values()));

    // Step 3: Iterate over each sheet group
    for (const [sheetName, sheetDataList] of groupedBySheet.entries()) {
      // Step 3.1: Read from sheet (headers and data area)
      console.log(`\nProcessing sheet: ${sheetName}`);
      const sheetData = await this.sheetsClient.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!B6:BB10`,
      });
      const values: string[][] = sheetData.data.values as string[][];
      const headerYearRow = values[0]; // Row 6
      const headerTypeRow = values[1]; // Row 7
      console.log('Header (Year/Month):', headerYearRow);
      console.log('Header (Type):', headerTypeRow);

      // Step 3.2: Build column map
      type ColumnInfo = { planCol?: string; factCol?: string };
      const columnMap = new Map<string, ColumnInfo>();
      for (let i = 0; i < headerTypeRow.length; i++) {
        const type = headerTypeRow[i]?.trim().toLowerCase();
        if (!type || (type !== 'п' && type !== 'ф')) continue;

        let yearMonth = '';
        for (let j = i; j >= 0; j--) {
          if (headerYearRow[j]?.trim()) {
            yearMonth = headerYearRow[j].trim();
            break;
          }
        }
        if (!yearMonth) continue;

        const [monthStr, yearStr] = yearMonth.split(/[.\s]+/);
        if (!monthStr || !yearStr) continue;

        const year = parseInt(yearStr);
        const monthIndex = this.monthStrToNumber(monthStr.toLowerCase());
        const key = `${year}-${monthIndex}`;
        const colLetter = this.indexToColumnLetter(i + 2);

        const existing: ColumnInfo = columnMap.get(key) || {};
        if (type === 'п') existing.planCol = colLetter;
        if (type === 'ф') existing.factCol = colLetter;

        columnMap.set(key, existing);
        console.log('Built column map:', columnMap);
      }

      // Step 3.3: Build row map
      const objectRows = values.slice(2); // Starting from row 8
      const rowMap = new Map<string, number>();
      for (let i = 0; i < objectRows.length; i++) {
        const objectName = objectRows[i][0]?.trim();
        if (objectName) {
          rowMap.set(objectName, i + 8); // +8 for actual row index
        }
      }
      console.log('Built row map:', rowMap);

      // Step 3.4: Prepare batch requests
      const requests: sheets_v4.Schema$ValueRange[] = [];
      for (const item of sheetDataList) {
        const { ObjectName, Year, Month, Plan, Fact } = item;
        const key = `${Year}-${Month}`;
        const colInfo = columnMap.get(key);
        const row = rowMap.get(ObjectName);

        if (!colInfo || !row) {
          console.log(`Skipping: ${ObjectName} ${key} — missing column or row`);
          continue;
        }
        console.log(
          `Updating: ${ObjectName} (${key}) — Plan: ${Plan}, Fact: ${Fact}`,
        );

        if (colInfo.planCol) {
          requests.push({
            range: `${sheetName}!${colInfo.planCol}${row}`,
            values: [[Plan]],
          });
        }
        if (colInfo.factCol) {
          requests.push({
            range: `${sheetName}!${colInfo.factCol}${row}`,
            values: [[Fact]],
          });
        }
      }

      // Step 3.5: Send batch update
      if (requests.length > 0) {
        console.log(`Sending batch update with ${requests.length} cells...`);
        console.log(
          'Sample of request payload:',
          JSON.stringify(requests.slice(0, 3), null, 2),
        );
        await this.sheetsClient.spreadsheets.values.batchUpdate({
          spreadsheetId,
          requestBody: {
            valueInputOption: 'USER_ENTERED',
            data: requests,
          },
        });
        console.log(`[${sheetName}] Updated ${requests.length} cells`);
      } else {
        console.log(`[${sheetName}] No updates to send`);
      }
    }
    return {
      message: 'Data successfully written to Google Sheet',
      updatedSheets: ['ВА', 'Б'],
    };
  }

  // Helper: Convert Russian month string to number
  private monthStrToNumber(month: string): number {
    const map: { [key: string]: number } = {
      янв: 1,
      февр: 2,
      мар: 3,
      апр: 4,
      мая: 5,
      июн: 6,
      июл: 7,
      авг: 8,
      сент: 9,
      окт: 10,
      нояб: 11,
      дек: 12,
    };
    return map[month.toLowerCase()] || 0;
  }

  // Helper: Convert column index to letter (e.g., 3 -> C)
  private indexToColumnLetter(index: number): string {
    let letter = '';
    while (index > 0) {
      const mod = (index - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      index = Math.floor((index - mod) / 26);
    }
    return letter;
  }

  // Helper: Convert column letter to index (e.g., C -> 3)
  // private columnLetterToIndex(letter: string): number {
  //   let index = 0;
  //   for (let i = 0; i < letter.length; i++) {
  //     index *= 26;
  //     index += letter.charCodeAt(i) - 64;
  //   }
  //   return index;
  // }
}
