WARNING: Could not find C:\temp\FAKTURA-1001_ADM-EIA-P.xlsx
Add-Worksheet : Cannot bind argument to parameter 'ExcelPackage' because it is null.
At C:\temp\Kristian\copilot_lic.ps1:114 char:39
+     $ws = Add-Worksheet -ExcelPackage $xl -WorksheetName "Faktura"
+                                       ~~~
    + CategoryInfo          : InvalidData: (:) [Add-Worksheet], ParameterBindingValidationException
    + FullyQualifiedErrorId : ParameterArgumentValidationErrorNullNotAllowed,Add-Worksheet
 
Cannot index into a null array.
At C:\temp\Kristian\copilot_lic.ps1:119 char:9
+         $sheet.Cells[$row, $col].Value = $value
+         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : NullArray
