import { Controller, Post } from '@nestjs/common';
import { GoogleSheetsService } from './google-sheets.service';

@Controller('google-sheets')
export class GoogleSheetsController {
  constructor(private readonly googleSheetsService: GoogleSheetsService) {}

  @Post('write')
  async writeDataToGoogleSheets() {
    return this.googleSheetsService.writeDataToGoogleSheets();
  }
}
