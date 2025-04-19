import { Controller, Get } from '@nestjs/common';
import { GoogleSheetsService } from './google-sheets.service';

@Controller('google-sheets')
export class GoogleSheetsController {
  constructor(private readonly googleSheetsService: GoogleSheetsService) {}

  @Get('write')
  async writeDataToGoogleSheets() {
    return this.googleSheetsService.writeDataToGoogleSheets();
  }
}
