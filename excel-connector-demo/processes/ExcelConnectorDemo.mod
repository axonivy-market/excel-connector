[Ivy]
17CFF194E102686A 9.3.0 #module
>Proto >Proto Collection #zClass
Eo0 ExcelConnectorDemo Big #zClass
Eo0 B #cInfo
Eo0 #process
Eo0 @AnnotationInP-0n ai ai #zField
Eo0 @TextInP .type .type #zField
Eo0 @TextInP .processKind .processKind #zField
Eo0 @TextInP .xml .xml #zField
Eo0 @TextInP .responsibility .responsibility #zField
Eo0 @StartRequest f0 '' #zField
Eo0 @EndTask f1 '' #zField
Eo0 @UserDialog f3 '' #zField
Eo0 @PushWFArc f4 '' #zField
Eo0 @PushWFArc f2 '' #zField
Eo0 @InfoButton f5 '' #zField
>Proto Eo0 Eo0 ExcelConnectorDemo #zField
Eo0 f0 outLink excelImportExport.ivp #txt
Eo0 f0 inParamDecl '<> param;' #txt
Eo0 f0 requestEnabled true #txt
Eo0 f0 triggerEnabled false #txt
Eo0 f0 callSignature excelImportExport() #txt
Eo0 f0 caseData businessCase.attach=true #txt
Eo0 f0 @CG|tags demo #txt
Eo0 f0 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>excelImportExport.ivp</name>
    </language>
</elementInfo>
' #txt
Eo0 f0 @C|.responsibility Everybody #txt
Eo0 f0 89 129 30 30 -21 17 #rect
Eo0 f1 377 129 30 30 0 15 #rect
Eo0 f3 dialogId com.axonivy.connector.excel.demo.PersonEditor #txt
Eo0 f3 startMethod start() #txt
Eo0 f3 requestActionDecl '<> param;' #txt
Eo0 f3 responseMappingAction 'out=in;
' #txt
Eo0 f3 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>PersonEditor</name>
    </language>
</elementInfo>
' #txt
Eo0 f3 192 122 112 44 -36 -8 #rect
Eo0 f4 119 144 192 144 #arcP
Eo0 f2 304 144 377 144 #arcP
Eo0 f5 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>This demo shows how easy it is to a) import data records from your local Excel files and convert them to Ivy Recordsets and&#13;
b) export your process data in the form of Ivy Recordsets as Excel files. Both .xls and .xlsx file types are supported.&#13;
For the import and export of excel files the two subprocesses ReadExcel and WriteExcel are used.</name>
    </language>
</elementInfo>
' #txt
Eo0 f5 24 26 688 60 -335 -24 #rect
>Proto Eo0 .type com.axonivy.connector.excel.demo.Data #txt
>Proto Eo0 .processKind NORMAL #txt
>Proto Eo0 0 0 32 24 18 0 #rect
>Proto Eo0 @|BIcon #fIcon
Eo0 f0 mainOut f4 tail #connect
Eo0 f4 head f3 mainIn #connect
Eo0 f3 mainOut f2 tail #connect
Eo0 f2 head f1 mainIn #connect
