import { Module } from '@nestjs/common';
import { GoogleSheetsService } from './google-sheets.service';
import { GoogleSheetsController } from './google-sheets.controller';
import { HttpModule } from '@nestjs/axios';

@Module({
  imports: [HttpModule],
  controllers: [GoogleSheetsController],
  providers: [GoogleSheetsService],
})
export class GoogleSheetsModule {}
