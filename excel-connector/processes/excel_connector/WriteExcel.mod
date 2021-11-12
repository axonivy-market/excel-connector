[Ivy]
17CF0992EAD1CDCB 9.3.0 #module
>Proto >Proto Collection #zClass
Wl0 WriteExcel Big #zClass
Wl0 B #cInfo
Wl0 #process
Wl0 @AnnotationInP-0n ai ai #zField
Wl0 @TextInP .type .type #zField
Wl0 @TextInP .processKind .processKind #zField
Wl0 @TextInP .xml .xml #zField
Wl0 @TextInP .responsibility .responsibility #zField
Wl0 @GridStep f2 '' #zField
Wl0 @StartSub f7 '' #zField
Wl0 @EndSub f3 '' #zField
Wl0 @PushWFArc f4 '' #zField
Wl0 @PushWFArc f0 '' #zField
Wl0 @StartSub f1 '' #zField
Wl0 @PushWFArc f5 '' #zField
>Proto Wl0 Wl0 WriteExcel #zField
Wl0 f2 actionTable 'out=in;
' #txt
Wl0 f2 actionCode 'import com.axonivy.connector.excel.ExcelWriter;

new ExcelWriter().execute(in.recordset, in.excelFile.getAbsolutePath());' #txt
Wl0 f2 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>call WriteExcel</name>
    </language>
</elementInfo>
' #txt
Wl0 f2 128 58 112 44 -39 -8 #rect
Wl0 f7 inParamDecl '<File excelFile,Recordset recordset> param;' #txt
Wl0 f7 inParamInfo 'excelFile.description=File path of the Excel sheet to be written.
recordset.description=Recordset represents the content of the Excel sheet to be written.' #txt
Wl0 f7 inParamTable 'out.excelFile=param.excelFile;
out.recordset=param.recordset;
' #txt
Wl0 f7 outParamDecl '<> result;' #txt
Wl0 f7 callSignature write(File,Recordset) #txt
Wl0 f7 @CG|tags connector #txt
Wl0 f7 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>call(WriteExcel(File,Recordset))</name>
    </language>
</elementInfo>
' #txt
Wl0 f7 25 65 30 30 -13 17 #rect
Wl0 f7 res:/webContent/icons/excel-icon.png?small #fDecoratorIcon
Wl0 f3 305 65 30 30 0 15 #rect
Wl0 f4 240 80 305 80 #arcP
Wl0 f0 55 80 128 80 #arcP
Wl0 f1 inParamDecl '<Recordset recordset> param;' #txt
Wl0 f1 inParamInfo 'recordset.description=Recordset represents the content of the Excel sheet to be written.' #txt
Wl0 f1 inParamTable 'out.excelFile=new File("export.xls",true);
out.recordset=param.recordset;
' #txt
Wl0 f1 outParamDecl '<File excelFile> result;' #txt
Wl0 f1 outParamInfo 'excelFile.description=Created Excel file.' #txt
Wl0 f1 outParamTable 'result.excelFile=in.excelFile;
' #txt
Wl0 f1 callSignature write(Recordset) #txt
Wl0 f1 @CG|tags connector #txt
Wl0 f1 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>call(WriteExcel(Recordset))</name>
    </language>
</elementInfo>
' #txt
Wl0 f1 33 185 30 30 -17 17 #rect
Wl0 f1 res:/webContent/icons/excel-icon.png?small #fDecoratorIcon
Wl0 f5 59 190 184 102 #arcP
>Proto Wl0 .type com.axonivy.connector.excel.WriteExcelData #txt
>Proto Wl0 .processKind CALLABLE_SUB #txt
>Proto Wl0 0 0 32 24 18 0 #rect
>Proto Wl0 @|BIcon #fIcon
Wl0 f2 mainOut f4 tail #connect
Wl0 f4 head f3 mainIn #connect
Wl0 f7 mainOut f0 tail #connect
Wl0 f0 head f2 mainIn #connect
Wl0 f1 mainOut f5 tail #connect
Wl0 f5 head f2 mainIn #connect
