/* Creëer een 'SEPA Pain direct debit core XML message' (pain.008.001.02) voor de ETT
* 
* 3 xml fragmenten worden gebruikt:
*  - root, Document en daarin CstmrDrctDbtInitn en de GrpHdr met alle content
*  - PmtInf, met de content muv. DrctDbtTxInf (de transactie informatie)
*  - DrctDbtTxInf, de transactie informatie
*
* Er zijn maximaal 2 PmtInf in het bericht: 1 voor de FRST, en 1 voor de RCUR
* >>> per eind 2016 ALLEEN NOG RCUR, dus nog maar 1 PmtInf
*  
* Per te incasseren lid wordt een DrctDbtTxInf aangemaakt en in de juiste PmtInf gehangen
*
* Inputparameters:
*  1. ddDate: gewenste incasso datum PmtInf/ReqdColltnDt formaat yyyy-mm-dd
*  2. debtors: lijst met debtor objecten; zie Tests.gs/testPainGen() voor voorbeeld objecten
*
* Algoritme:
*  1. parseer de 3 xml berichten;
*  2. per incasso regel (itereer over debtors):
*     a. maak de DrctDbtTxInf transactie informatie aan
*     b. koppel het aan PmtInf
*     c. werk de telling voor de NbOfTxs en CtrlSum velden in GrpHdr en PmtInf bij
*  4. als (aantal transacties in PmtInf > 0) dan { koppel PmtInf aan root };
*  5. werk de NbOfTxs en CtrlSum velden in GrpHdr en PmtInf bij
*  
*/

function createPainMessage(incSheet, debtors) {
  
  var ddDate = Utilities.formatDate(incSheet.getRange("datumIncasso").getValue(), "CET", "yyyy-MM-dd");
  var ddId = incSheet.getRange("kenmerkIncasso").getValue();
  
  // initialisaties
  var ns = XmlService.getNamespace("urn:iso:std:iso:20022:tech:xsd:pain.008.001.02");
  
  // parseer de 2 xml berichten, zet de derde klaar (wordt in de loop geparseerd)
  var rootText = "<Document xmlns='urn:iso:std:iso:20022:tech:xsd:pain.008.001.02' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'><CstmrDrctDbtInitn><GrpHdr><MsgId></MsgId><CreDtTm>2013-12-25T10:54:35</CreDtTm><NbOfTxs>1</NbOfTxs><CtrlSum>5.00</CtrlSum><InitgPty><Nm>Eerste Twellose Toerclub</Nm></InitgPty></GrpHdr></CstmrDrctDbtInitn></Document>";
  var rootDocument = XmlService.parse(rootText);
  var root = rootDocument.getRootElement();
  var CstmrDrctDbtInitn = root.getChild("CstmrDrctDbtInitn",ns);
  
  var PmtInfText = "<PmtInf xmlns='urn:iso:std:iso:20022:tech:xsd:pain.008.001.02'><PmtInfId>ETT-20131225-testv1</PmtInfId><PmtMtd>DD</PmtMtd><NbOfTxs>1</NbOfTxs><CtrlSum>5.00</CtrlSum><PmtTpInf><SvcLvl><Cd>SEPA</Cd></SvcLvl><LclInstrm><Cd>CORE</Cd></LclInstrm><SeqTp>RCUR</SeqTp></PmtTpInf><ReqdColltnDt>2014-01-02</ReqdColltnDt><Cdtr><Nm>Eerste Twellose Toerclub</Nm><PstlAdr><Ctry>NL</Ctry></PstlAdr></Cdtr><CdtrAcct><Id><IBAN>NL09RABO0111019788</IBAN></Id></CdtrAcct><CdtrAgt><FinInstnId><BIC>RABONL2U</BIC></FinInstnId></CdtrAgt><UltmtCdtr><Nm>Eerste Twellose Toerclub</Nm></UltmtCdtr><ChrgBr>SLEV</ChrgBr><CdtrSchmeId><Id><PrvtId><Othr><Id>NL97ZZZ401029480000</Id><SchmeNm><Prtry>SEPA</Prtry></SchmeNm></Othr></PrvtId></Id></CdtrSchmeId></PmtInf>";
  var PmtInf  = XmlService.parse(PmtInfText).detachRootElement();
  
  var DrctDbtTxInfText = "<DrctDbtTxInf xmlns='urn:iso:std:iso:20022:tech:xsd:pain.008.001.02'><PmtId><EndToEndId></EndToEndId></PmtId><InstdAmt Ccy='EUR'>${'instdamt'}</InstdAmt><DrctDbtTx><MndtRltdInf><MndtId>${'mndtid'}</MndtId><DtOfSgntr>${'dtofsgntr'}</DtOfSgntr><AmdmntInd>${'amdmntind'}</AmdmntInd></MndtRltdInf></DrctDbtTx><DbtrAgt><FinInstnId/></DbtrAgt><Dbtr><Nm>${'dbtrnm'}</Nm><PstlAdr><Ctry>NL</Ctry></PstlAdr></Dbtr><DbtrAcct><Id><IBAN>${'dbtriban'}</IBAN></Id></DbtrAcct><UltmtDbtr><Nm>${'dbtrnm'}</Nm></UltmtDbtr><RmtInf><Ustrd>${'dbtrustrd'}</Ustrd></RmtInf></DrctDbtTxInf>";
  
  // per incasso regel (itereer over debtors) en hou aantal transacties (NbOfTxs) en controlesom (CtrlSum) bij:
  var NbOfTxs = 0;
  var CtrlSum = 0;
  
  for (var i = 0; i < debtors.length; i++) {
    if (debtors[i]["betaling"] == "INCASSO" || debtors[i]["betaling"] == "SEC-INC") {
      var DrctDbtTxInf = XmlService.parse(DrctDbtTxInfText).detachRootElement();
      
      setXmlElementValue(DrctDbtTxInf, 'PmtId/EndToEndId', ddId + "-" + debtors[i]["mndtid"], ns);
      setXmlElementValue(DrctDbtTxInf, 'InstdAmt', debtors[i]["instdamt"], ns);
      setXmlElementValue(DrctDbtTxInf, 'DrctDbtTx/MndtRltdInf/MndtId', debtors[i]["mndtid"], ns);
      setXmlElementValue(DrctDbtTxInf, 'DrctDbtTx/MndtRltdInf/DtOfSgntr', debtors[i]["dtofsgntr"], ns);
      setXmlElementValue(DrctDbtTxInf, 'DrctDbtTx/MndtRltdInf/AmdmntInd', debtors[i]["amdmntind"], ns);
      setXmlElementValue(DrctDbtTxInf, 'Dbtr/Nm', debtors[i]["dbtrnm"], ns);
      setXmlElementValue(DrctDbtTxInf, 'UltmtDbtr/Nm', debtors[i]["dbtrnm"], ns);
      setXmlElementValue(DrctDbtTxInf, 'DbtrAcct/Id/IBAN', debtors[i]["dbtriban"], ns);
      setXmlElementValue(DrctDbtTxInf, 'RmtInf/Ustrd', debtors[i]["dbtrustrd"], ns);
      
      NbOfTxs += 1;
      CtrlSum += parseFloat(debtors[i]["instdamt"]);
      PmtInf.addContent(DrctDbtTxInf);
    }
  }
  
  // GrpHdr
  setXmlElementValue(CstmrDrctDbtInitn, 'GrpHdr/MsgId', ddId+Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd'T'HH:MM:ss"), ns);
  setXmlElementValue(CstmrDrctDbtInitn, 'GrpHdr/CreDtTm', Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd'T'HH:MM:ss"), ns);
  setXmlElementValue(CstmrDrctDbtInitn, 'GrpHdr/NbOfTxs', Utilities.formatString('%d', (NbOfTxs)), ns);
  setXmlElementValue(CstmrDrctDbtInitn, 'GrpHdr/CtrlSum', Utilities.formatString('%.2f', (CtrlSum)), ns);
  
  // PmtInf
  if (NbOfTxs > 0){
    setXmlElementValue(PmtInf, 'PmtInfId', ddId +'-RCUR', ns);
    setXmlElementValue(PmtInf, 'NbOfTxs', Utilities.formatString('%d', NbOfTxs), ns);
    setXmlElementValue(PmtInf, 'CtrlSum', Utilities.formatString('%.2f', CtrlSum), ns);
    setXmlElementValue(PmtInf, 'ReqdColltnDt', ddDate, ns);
    
    CstmrDrctDbtInitn.addContent(PmtInf);
  };
  
  var xml = XmlService.getPrettyFormat()
  .setLineSeparator('\n')
  .format(rootDocument);
  
  var fileName = ddId + ' SepaDD' + ddDate + 'v1.xml';
  var file = DriveApp.createFile(Utilities.newBlob(xml, MimeType.PLAIN_TEXT, fileName));
  file.setName(fileName);
  
  Logger.log(fileName);
  
  /*
  painDoc = DocumentApp.create("pain.xml");
  painDoc.getBody().setText(xml);
  
  Logger.log("vlak voor het saven");
  painDoc.saveAndClose();
  Logger.log("vlak na het saven");
  */
}

// hulpfunctie om in een xml fragment de waarde van een element aan te passen
// let op: fragment wordt (denk ik ;-) 'by reference' doorgegeven
function setXmlElementValue(fragment, path, value, ns) {
  var descendant = fragment;
  var pathElement = path.split('/');
  for (var i in pathElement) {
    descendant = descendant.getChild(pathElement[i], ns);
  }
  descendant.setText(value);
  return fragment;
}

function compare(a,b) {
  if (a.mndtid < b.mndtid)
    return -1;
  if (a.mndtid > b.mndtid)
    return 1;
  return 0;
}