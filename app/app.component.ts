import { Component } from '@angular/core';
import { aggregateBy, process } from '@progress/kendo-data-query';
import { products } from './products';

@Component({
  selector: 'my-app',
  template: `
        <button type="button" class="k-button" (click)="save(excelexport, excelexport1)">Export To Excel</button>

        <kendo-excelexport [data]="data" fileName="Products.xlsx" #excelexport>
            <kendo-excelexport-column field="ProductID" [locked]="true" title="Product ID" [width]="200">
            </kendo-excelexport-column>
            <kendo-excelexport-column field="ProductName" title="Product Name" [width]="350">
            </kendo-excelexport-column>
            <kendo-excelexport-column field="UnitPrice" title="Unit Price" [width]="120">
            </kendo-excelexport-column>
            <kendo-excelexport-column field="Seat" title="WorkStation" [width]="120" [cellOptions]="timeOptions">
            </kendo-excelexport-column>
      </kendo-excelexport>
        <kendo-excelexport [data]="data1" fileName="Products.xlsx" #excelexport1>
            <kendo-excelexport-column field="ProductName" title="Product Name" [width]="50">
            </kendo-excelexport-column>
      </kendo-excelexport>
    `,
})
export class AppComponent {
  public data: any[] = products.slice(0, 40);
  public data1: any[] = products.slice(40);

  timeOptions = {
    validation: {
      dataType: 'list',
      showButton: true,
      comparerType: 'list',
      allowNulls: true,
      type: 'reject',
      from: 'Sheet2!A$2:A$' + this.data1.length,
    },
  };

  public save(component1, component2): void {
    Promise.all([
      component1.workbookOptions(),
      component2.workbookOptions(),
    ]).then((workbooks) => {
      workbooks[0].sheets = workbooks[0].sheets.concat(workbooks[1].sheets);
      component1.save(workbooks[0]);
    });
  }
}
