import { BaseCommand } from '@adonisjs/core/build/standalone'
import File from 'App/Models/File'
import ExcelJS from 'exceljs';
import Client from 'App/Models/Client'
import Bank from 'App/Models/Bank'
import Address from 'App/Models/Address'
import { DateTime } from 'luxon'

interface BanksInterfaces {
  id:number,
  clientId:number,
  bankName:string,
  accountHolderName: string,
  accountNumber:string,
  ifscCode:string | undefined,
  address:string | undefined,
  city:string | undefined,
  createdAt:DateTime,
  updatedAt:DateTime,
}
interface AddressInterfaces {
  id: number;
  clientId: number;
  addressLine1: string;
  addressLine2: string | undefined;
  city: string;
  state: string | undefined;
  zip: string | undefined;  
  createdAt: DateTime;
  updatedAt: DateTime;
}

interface ClientInterfaces {
  id: number;
  name: string;
  email: string | undefined;
  phoneNumber: string | undefined;
  pan: string | undefined;
  createdAt: DateTime;
  updatedAt: DateTime;
}


export default class ProcessFile extends BaseCommand {
  public static commandName = 'process:file';

  public static description = 'process excel file into database';

  public static settings = {
    loadApp: true,
  };

  public async run() {
    const files = await File.query().orderBy('id').limit(1);
    this.processFile(files);
  }

  public async processFile(files: File[]) {
    const file = files[0];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file.filePath);

    // Process Client Sheet (Sheet 1)
    const clientSheet = workbook.getWorksheet('Sheet1');
    clientSheet.eachRow(async (row:any) => {
      const clientRow: ClientInterfaces = {
        // Map row values to the ClientInterface properties
        id: row.getCell(1).value,
        name: row.getCell(2).value,
        email: row.getCell(3).value,
        phoneNumber: row.getCell(4).value,
        pan: row.getCell(5).value,
        createdAt: DateTime.fromJSDate(row.getCell(6).value),
        updatedAt: DateTime.fromJSDate(row.getCell(7).value),
      };
      let ans=await Client.create(clientRow);
      console.log(ans)
    });

    // Process Bank Detail Sheet (Sheet 2)
    const bankDetailSheet = workbook.getWorksheet('Sheet2');
    bankDetailSheet.eachRow(async (row:any) => {
      const bankDetailRow: BanksInterfaces = {
        id: row.getCell(1).value,
        clientId: row.getCell(2).value,
        bankName: row.getCell(3).value,
        accountHolderName: row.getCell(4).value,
        accountNumber:row.getCell(5).value,
        ifscCode: row.getCell(6).value ,
        address: row.getCell(7).value,
        city: row.getCell(8).value,
        createdAt: DateTime.fromJSDate(row.getCell(9).value),
        updatedAt: DateTime.fromJSDate(row.getCell(10).value),
      };
     let ans= await Bank.create(bankDetailRow);
     console.log(ans)
    });

    // Process Address Sheet (Sheet 3)
    const addressSheet = workbook.getWorksheet('Sheet3');
    addressSheet.eachRow(async (row:any) => {
      const addressRow: AddressInterfaces = {
        // Map row values to the AddressInterface properties
        id:row.getCell(1).value,
        clientId:row.getCell(2).value,
        addressLine1:row.getCell(3).value,
        addressLine2:row.getCell(4).value,
        city:row.getCell(5).value,
        state:row.getCell(6).value,
        zip:row.getcell(7).value,
        createdAt:DateTime.fromJSDate(row.getCell(8).value),
        updatedAt:DateTime.fromJSDate(row.getCell(9).value),
      };
     let ans= await Address.create(addressRow);
     console.log(ans)
    });

    console.log('Excel file processing complete.');
  }
}