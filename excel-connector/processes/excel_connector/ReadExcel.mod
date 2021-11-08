[Ivy]
17CF0967FCDFADEC 9.3.0 #module
>Proto >Proto Collection #zClass
Rl0 ReadExcel Big #zClass
Rl0 B #cInfo
Rl0 #process
Rl0 @AnnotationInP-0n ai ai #zField
Rl0 @TextInP .type .type #zField
Rl0 @TextInP .processKind .processKind #zField
Rl0 @TextInP .xml .xml #zField
Rl0 @TextInP .responsibility .responsibility #zField
Rl0 @GridStep f1 '' #zField
Rl0 @StartSub f7 '' #zField
Rl0 @EndSub f3 '' #zField
Rl0 @PushWFArc f5 '' #zField
Rl0 @PushWFArc f0 '' #zField
>Proto Rl0 Rl0 ReadExcel #zField
Rl0 f1 actionTable 'out=in;
' #txt
Rl0 f1 actionCode 'import com.axonivy.connector.excel.ExcelReader;

out.recordset = new ExcelReader().getRecordset(in.filePath);' #txt
Rl0 f1 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>callReadExcel</name>
    </language>
</elementInfo>
' #txt
Rl0 f1 176 42 112 44 -39 -8 #rect
Rl0 f7 inParamDecl '<String filePath> param;' #txt
Rl0 f7 inParamInfo 'filePath.description=File path of the Excel sheet to be read.' #txt
Rl0 f7 inParamTable 'out.filePath=param.filePath;
' #txt
Rl0 f7 outParamDecl '<> result;' #txt
Rl0 f7 callSignature call(String) #txt
Rl0 f7 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>call(String)</name>
    </language>
</elementInfo>
' #txt
Rl0 f7 73 49 30 30 -13 17 #rect
Rl0 f7 res:/webContent/icons/excel-icon.png?small #fDecoratorIcon
Rl0 f3 361 49 30 30 0 15 #rect
Rl0 f5 288 64 361 64 #arcP
Rl0 f0 103 64 176 64 #arcP
>Proto Rl0 .type com.axonivy.connector.excel.ReadExcelData #txt
>Proto Rl0 .processKind CALLABLE_SUB #txt
>Proto Rl0 0 0 32 24 18 0 #rect
>Proto Rl0 @|BIcon #fIcon
Rl0 f1 mainOut f5 tail #connect
Rl0 f5 head f3 mainIn #connect
Rl0 f7 mainOut f0 tail #connect
Rl0 f0 head f1 mainIn #connect
