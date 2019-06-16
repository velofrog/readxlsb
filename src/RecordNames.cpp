#include "RecordNames.h"

namespace readxlsb
{

std::string TokenLabel(uint32_t record_id) {
    switch (record_id) {
    case 0:
        return "BrtRowHdr";
        break;
    case 1:
        return "BrtCellBlank";
        break;
    case 2:
        return "BrtCellRk";
        break;
    case 3:
        return "BrtCellError";
        break;
    case 4:
        return "BrtCellBool";
        break;
    case 5:
        return "BrtCellReal";
        break;
    case 6:
        return "BrtCellSt";
        break;
    case 7:
        return "BrtCellIsst";
        break;
    case 8:
        return "BrtFmlaString";
        break;
    case 9:
        return "BrtFmlaNum";
        break;
    case 10:
        return "BrtFmlaBool";
        break;
    case 11:
        return "BrtFmlaError";
        break;
    case 19:
        return "BrtSSTItem";
        break;
    case 20:
        return "BrtPCDIMissing";
        break;
    case 21:
        return "BrtPCDINumber";
        break;
    case 22:
        return "BrtPCDIBoolean";
        break;
    case 23:
        return "BrtPCDIError";
        break;
    case 24:
        return "BrtPCDIString";
        break;
    case 25:
        return "BrtPCDIDatetime";
        break;
    case 26:
        return "BrtPCDIIndex";
        break;
    case 27:
        return "BrtPCDIAMissing";
        break;
    case 28:
        return "BrtPCDIANumber";
        break;
    case 29:
        return "BrtPCDIABoolean";
        break;
    case 30:
        return "BrtPCDIAError";
        break;
    case 31:
        return "BrtPCDIAString";
        break;
    case 32:
        return "BrtPCDIADatetime";
        break;
    case 33:
        return "BrtPCRRecord";
        break;
    case 34:
        return "BrtPCRRecordDt";
        break;
    case 35:
        return "BrtFRTBegin";
        break;
    case 36:
        return "BrtFRTEnd";
        break;
    case 37:
        return "BrtACBegin";
        break;
    case 38:
        return "BrtACEnd";
        break;
    case 39:
        return "BrtName";
        break;
    case 40:
        return "BrtIndexRowBlock";
        break;
    case 42:
        return "BrtIndexBlock";
        break;
    case 43:
        return "BrtFont";
        break;
    case 44:
        return "BrtFmt";
        break;
    case 45:
        return "BrtFill";
        break;
    case 46:
        return "BrtBorder";
        break;
    case 47:
        return "BrtXF";
        break;
    case 48:
        return "BrtStyle";
        break;
    case 49:
        return "BrtCellMeta";
        break;
    case 50:
        return "BrtValueMeta";
        break;
    case 51:
        return "BrtMdb";
        break;
    case 52:
        return "BrtBeginFmd";
        break;
    case 53:
        return "BrtEndFmd";
        break;
    case 54:
        return "BrtBeginMdx";
        break;
    case 55:
        return "BrtEndMdx";
        break;
    case 56:
        return "BrtBeginMdxTuple";
        break;
    case 57:
        return "BrtEndMdxTuple";
        break;
    case 58:
        return "BrtMdxMbrIstr";
        break;
    case 59:
        return "BrtStr";
        break;
    case 60:
        return "BrtColInfo";
        break;
    case 62:
        return "BrtCellRString";
        break;
    case 64:
        return "BrtDVal";
        break;
    case 65:
        return "BrtSxvcellNum";
        break;
    case 66:
        return "BrtSxvcellStr";
        break;
    case 67:
        return "BrtSxvcellBool";
        break;
    case 68:
        return "BrtSxvcellErr";
        break;
    case 69:
        return "BrtSxvcellDate";
        break;
    case 70:
        return "BrtSxvcellNil";
        break;
    case 128:
        return "BrtFileVersion";
        break;
    case 129:
        return "BrtBeginSheet";
        break;
    case 130:
        return "BrtEndSheet";
        break;
    case 131:
        return "BrtBeginBook";
        break;
    case 132:
        return "BrtEndBook";
        break;
    case 133:
        return "BrtBeginWsViews";
        break;
    case 134:
        return "BrtEndWsViews";
        break;
    case 135:
        return "BrtBeginBookViews";
        break;
    case 136:
        return "BrtEndBookViews";
        break;
    case 137:
        return "BrtBeginWsView";
        break;
    case 138:
        return "BrtEndWsView";
        break;
    case 139:
        return "BrtBeginCsViews";
        break;
    case 140:
        return "BrtEndCsViews";
        break;
    case 141:
        return "BrtBeginCsView";
        break;
    case 142:
        return "BrtEndCsView";
        break;
    case 143:
        return "BrtBeginBundleShs";
        break;
    case 144:
        return "BrtEndBundleShs";
        break;
    case 145:
        return "BrtBeginSheetData";
        break;
    case 146:
        return "BrtEndSheetData";
        break;
    case 147:
        return "BrtWsProp";
        break;
    case 148:
        return "BrtWsDim";
        break;
    case 151:
        return "BrtPane";
        break;
    case 152:
        return "BrtSel";
        break;
    case 153:
        return "BrtWbProp";
        break;
    case 154:
        return "BrtWbFactoid";
        break;
    case 155:
        return "BrtFileRecover";
        break;
    case 156:
        return "BrtBundleSh";
        break;
    case 157:
        return "BrtCalcProp";
        break;
    case 158:
        return "BrtBookView";
        break;
    case 159:
        return "BrtBeginSst";
        break;
    case 160:
        return "BrtEndSst";
        break;
    case 161:
        return "BrtBeginAFilter";
        break;
    case 162:
        return "BrtEndAFilter";
        break;
    case 163:
        return "BrtBeginFilterColumn";
        break;
    case 164:
        return "BrtEndFilterColumn";
        break;
    case 165:
        return "BrtBeginFilters";
        break;
    case 166:
        return "BrtEndFilters";
        break;
    case 167:
        return "BrtFilter";
        break;
    case 168:
        return "BrtColorFilter";
        break;
    case 169:
        return "BrtIconFilter";
        break;
    case 170:
        return "BrtTop10Filter";
        break;
    case 171:
        return "BrtDynamicFilter";
        break;
    case 172:
        return "BrtBeginCustomFilters";
        break;
    case 173:
        return "BrtEndCustomFilters";
        break;
    case 174:
        return "BrtCustomFilter";
        break;
    case 175:
        return "BrtAFilterDateGroupItem";
        break;
    case 176:
        return "BrtMergeCell";
        break;
    case 177:
        return "BrtBeginMergeCells";
        break;
    case 178:
        return "BrtEndMergeCells";
        break;
    case 179:
        return "BrtBeginPivotCacheDef";
        break;
    case 180:
        return "BrtEndPivotCacheDef";
        break;
    case 181:
        return "BrtBeginPCDFields";
        break;
    case 182:
        return "BrtEndPCDFields";
        break;
    case 183:
        return "BrtBeginPCDField";
        break;
    case 184:
        return "BrtEndPCDField";
        break;
    case 185:
        return "BrtBeginPCDSource";
        break;
    case 186:
        return "BrtEndPCDSource";
        break;
    case 187:
        return "BrtBeginPCDSRange";
        break;
    case 188:
        return "BrtEndPCDSRange";
        break;
    case 189:
        return "BrtBeginPCDFAtbl";
        break;
    case 190:
        return "BrtEndPCDFAtbl";
        break;
    case 191:
        return "BrtBeginPCDIRun";
        break;
    case 192:
        return "BrtEndPCDIRun";
        break;
    case 193:
        return "BrtBeginPivotCacheRecords";
        break;
    case 194:
        return "BrtEndPivotCacheRecords";
        break;
    case 195:
        return "BrtBeginPCDHierarchies";
        break;
    case 196:
        return "BrtEndPCDHierarchies";
        break;
    case 197:
        return "BrtBeginPCDHierarchy";
        break;
    case 198:
        return "BrtEndPCDHierarchy";
        break;
    case 199:
        return "BrtBeginPCDHFieldsUsage";
        break;
    case 200:
        return "BrtEndPCDHFieldsUsage";
        break;
    case 201:
        return "BrtBeginExtConnection";
        break;
    case 202:
        return "BrtEndExtConnection";
        break;
    case 203:
        return "BrtBeginECDbProps";
        break;
    case 204:
        return "BrtEndECDbProps";
        break;
    case 205:
        return "BrtBeginECOlapProps";
        break;
    case 206:
        return "BrtEndECOlapProps";
        break;
    case 207:
        return "BrtBeginPCDSConsol";
        break;
    case 208:
        return "BrtEndPCDSConsol";
        break;
    case 209:
        return "BrtBeginPCDSCPages";
        break;
    case 210:
        return "BrtEndPCDSCPages";
        break;
    case 211:
        return "BrtBeginPCDSCPage";
        break;
    case 212:
        return "BrtEndPCDSCPage";
        break;
    case 213:
        return "BrtBeginPCDSCPItem";
        break;
    case 214:
        return "BrtEndPCDSCPItem";
        break;
    case 215:
        return "BrtBeginPCDSCSets";
        break;
    case 216:
        return "BrtEndPCDSCSets";
        break;
    case 217:
        return "BrtBeginPCDSCSet";
        break;
    case 218:
        return "BrtEndPCDSCSet";
        break;
    case 219:
        return "BrtBeginPCDFGroup";
        break;
    case 220:
        return "BrtEndPCDFGroup";
        break;
    case 221:
        return "BrtBeginPCDFGItems";
        break;
    case 222:
        return "BrtEndPCDFGItems";
        break;
    case 223:
        return "BrtBeginPCDFGRange";
        break;
    case 224:
        return "BrtEndPCDFGRange";
        break;
    case 225:
        return "BrtBeginPCDFGDiscrete";
        break;
    case 226:
        return "BrtEndPCDFGDiscrete";
        break;
    case 227:
        return "BrtBeginPCDSDTupleCache";
        break;
    case 228:
        return "BrtEndPCDSDTupleCache";
        break;
    case 229:
        return "BrtBeginPCDSDTCEntries";
        break;
    case 230:
        return "BrtEndPCDSDTCEntries";
        break;
    case 231:
        return "BrtBeginPCDSDTCEMembers";
        break;
    case 232:
        return "BrtEndPCDSDTCEMembers";
        break;
    case 233:
        return "BrtBeginPCDSDTCEMember";
        break;
    case 234:
        return "BrtEndPCDSDTCEMember";
        break;
    case 235:
        return "BrtBeginPCDSDTCQueries";
        break;
    case 236:
        return "BrtEndPCDSDTCQueries";
        break;
    case 237:
        return "BrtBeginPCDSDTCQuery";
        break;
    case 238:
        return "BrtEndPCDSDTCQuery";
        break;
    case 239:
        return "BrtBeginPCDSDTCSets";
        break;
    case 240:
        return "BrtEndPCDSDTCSets";
        break;
    case 241:
        return "BrtBeginPCDSDTCSet";
        break;
    case 242:
        return "BrtEndPCDSDTCSet";
        break;
    case 243:
        return "BrtBeginPCDCalcItems";
        break;
    case 244:
        return "BrtEndPCDCalcItems";
        break;
    case 245:
        return "BrtBeginPCDCalcItem";
        break;
    case 246:
        return "BrtEndPCDCalcItem";
        break;
    case 247:
        return "BrtBeginPRule";
        break;
    case 248:
        return "BrtEndPRule";
        break;
    case 249:
        return "BrtBeginPRFilters";
        break;
    case 250:
        return "BrtEndPRFilters";
        break;
    case 251:
        return "BrtBeginPRFilter";
        break;
    case 252:
        return "BrtEndPRFilter";
        break;
    case 253:
        return "BrtBeginPNames";
        break;
    case 254:
        return "BrtEndPNames";
        break;
    case 255:
        return "BrtBeginPName";
        break;
    case 256:
        return "BrtEndPName";
        break;
    case 257:
        return "BrtBeginPNPairs";
        break;
    case 258:
        return "BrtEndPNPairs";
        break;
    case 259:
        return "BrtBeginPNPair";
        break;
    case 260:
        return "BrtEndPNPair";
        break;
    case 261:
        return "BrtBeginECWebProps";
        break;
    case 262:
        return "BrtEndECWebProps";
        break;
    case 263:
        return "BrtBeginEcWpTables";
        break;
    case 264:
        return "BrtEndECWPTables";
        break;
    case 265:
        return "BrtBeginECParams";
        break;
    case 266:
        return "BrtEndECParams";
        break;
    case 267:
        return "BrtBeginECParam";
        break;
    case 268:
        return "BrtEndECParam";
        break;
    case 269:
        return "BrtBeginPCDKPIs";
        break;
    case 270:
        return "BrtEndPCDKPIs";
        break;
    case 271:
        return "BrtBeginPCDKPI";
        break;
    case 272:
        return "BrtEndPCDKPI";
        break;
    case 273:
        return "BrtBeginDims";
        break;
    case 274:
        return "BrtEndDims";
        break;
    case 275:
        return "BrtBeginDim";
        break;
    case 276:
        return "BrtEndDim";
        break;
    case 277:
        return "BrtIndexPartEnd";
        break;
    case 278:
        return "BrtBeginStyleSheet";
        break;
    case 279:
        return "BrtEndStyleSheet";
        break;
    case 280:
        return "BrtBeginSXView";
        break;
    case 281:
        return "BrtEndSXVI";
        break;
    case 282:
        return "BrtBeginSXVI";
        break;
    case 283:
        return "BrtBeginSXVIs";
        break;
    case 284:
        return "BrtEndSXVIs";
        break;
    case 285:
        return "BrtBeginSXVD";
        break;
    case 286:
        return "BrtEndSXVD";
        break;
    case 287:
        return "BrtBeginSXVDs";
        break;
    case 288:
        return "BrtEndSXVDs";
        break;
    case 289:
        return "BrtBeginSXPI";
        break;
    case 290:
        return "BrtEndSXPI";
        break;
    case 291:
        return "BrtBeginSXPIs";
        break;
    case 292:
        return "BrtEndSXPIs";
        break;
    case 293:
        return "BrtBeginSXDI";
        break;
    case 294:
        return "BrtEndSXDI";
        break;
    case 295:
        return "BrtBeginSXDIs";
        break;
    case 296:
        return "BrtEndSXDIs";
        break;
    case 297:
        return "BrtBeginSXLI";
        break;
    case 298:
        return "BrtEndSXLI";
        break;
    case 299:
        return "BrtBeginSXLIRws";
        break;
    case 300:
        return "BrtEndSXLIRws";
        break;
    case 301:
        return "BrtBeginSXLICols";
        break;
    case 302:
        return "BrtEndSXLICols";
        break;
    case 303:
        return "BrtBeginSXFormat";
        break;
    case 304:
        return "BrtEndSXFormat";
        break;
    case 305:
        return "BrtBeginSXFormats";
        break;
    case 306:
        return "BrtEndSxFormats";
        break;
    case 307:
        return "BrtBeginSxSelect";
        break;
    case 308:
        return "BrtEndSxSelect";
        break;
    case 309:
        return "BrtBeginISXVDRws";
        break;
    case 310:
        return "BrtEndISXVDRws";
        break;
    case 311:
        return "BrtBeginISXVDCols";
        break;
    case 312:
        return "BrtEndISXVDCols";
        break;
    case 313:
        return "BrtEndSXLocation";
        break;
    case 314:
        return "BrtBeginSXLocation";
        break;
    case 315:
        return "BrtEndSXView";
        break;
    case 316:
        return "BrtBeginSXTHs";
        break;
    case 317:
        return "BrtEndSXTHs";
        break;
    case 318:
        return "BrtBeginSXTH";
        break;
    case 319:
        return "BrtEndSXTH";
        break;
    case 320:
        return "BrtBeginISXTHRws";
        break;
    case 321:
        return "BrtEndISXTHRws";
        break;
    case 322:
        return "BrtBeginISXTHCols";
        break;
    case 323:
        return "BrtEndISXTHCols";
        break;
    case 324:
        return "BrtBeginSXTDMPS";
        break;
    case 325:
        return "BrtEndSXTDMPs";
        break;
    case 326:
        return "BrtBeginSXTDMP";
        break;
    case 327:
        return "BrtEndSXTDMP";
        break;
    case 328:
        return "BrtBeginSXTHItems";
        break;
    case 329:
        return "BrtEndSXTHItems";
        break;
    case 330:
        return "BrtBeginSXTHItem";
        break;
    case 331:
        return "BrtEndSXTHItem";
        break;
    case 332:
        return "BrtBeginMetadata";
        break;
    case 333:
        return "BrtEndMetadata";
        break;
    case 334:
        return "BrtBeginEsmdtinfo";
        break;
    case 335:
        return "BrtMdtinfo";
        break;
    case 336:
        return "BrtEndEsmdtinfo";
        break;
    case 337:
        return "BrtBeginEsmdb";
        break;
    case 338:
        return "BrtEndEsmdb";
        break;
    case 339:
        return "BrtBeginEsfmd";
        break;
    case 340:
        return "BrtEndEsfmd";
        break;
    case 341:
        return "BrtBeginSingleCells";
        break;
    case 342:
        return "BrtEndSingleCells";
        break;
    case 343:
        return "BrtBeginList";
        break;
    case 344:
        return "BrtEndList";
        break;
    case 345:
        return "BrtBeginListCols";
        break;
    case 346:
        return "BrtEndListCols";
        break;
    case 347:
        return "BrtBeginListCol";
        break;
    case 348:
        return "BrtEndListCol";
        break;
    case 349:
        return "BrtBeginListXmlCPr";
        break;
    case 350:
        return "BrtEndListXmlCPr";
        break;
    case 351:
        return "BrtListCCFmla";
        break;
    case 352:
        return "BrtListTrFmla";
        break;
    case 353:
        return "BrtBeginExternals";
        break;
    case 354:
        return "BrtEndExternals";
        break;
    case 355:
        return "BrtSupBookSrc";
        break;
    case 357:
        return "BrtSupSelf";
        break;
    case 358:
        return "BrtSupSame";
        break;
    case 359:
        return "BrtSupTabs";
        break;
    case 360:
        return "BrtBeginSupBook";
        break;
    case 361:
        return "BrtPlaceholderName";
        break;
    case 362:
        return "BrtExternSheet";
        break;
    case 363:
        return "BrtExternTableStart";
        break;
    case 364:
        return "BrtExternTableEnd";
        break;
    case 366:
        return "BrtExternRowHdr";
        break;
    case 367:
        return "BrtExternCellBlank";
        break;
    case 368:
        return "BrtExternCellReal";
        break;
    case 369:
        return "BrtExternCellBool";
        break;
    case 370:
        return "BrtExternCellError";
        break;
    case 371:
        return "BrtExternCellString";
        break;
    case 372:
        return "BrtBeginEsmdx";
        break;
    case 373:
        return "BrtEndEsmdx";
        break;
    case 374:
        return "BrtBeginMdxSet";
        break;
    case 375:
        return "BrtEndMdxSet";
        break;
    case 376:
        return "BrtBeginMdxMbrProp";
        break;
    case 377:
        return "BrtEndMdxMbrProp";
        break;
    case 378:
        return "BrtBeginMdxKPI";
        break;
    case 379:
        return "BrtEndMdxKPI";
        break;
    case 380:
        return "BrtBeginEsstr";
        break;
    case 381:
        return "BrtEndEsstr";
        break;
    case 382:
        return "BrtBeginPRFItem";
        break;
    case 383:
        return "BrtEndPRFItem";
        break;
    case 384:
        return "BrtBeginPivotCacheIDs";
        break;
    case 385:
        return "BrtEndPivotCacheIDs";
        break;
    case 386:
        return "BrtBeginPivotCacheID";
        break;
    case 387:
        return "BrtEndPivotCacheID";
        break;
    case 388:
        return "BrtBeginISXVIs";
        break;
    case 389:
        return "BrtEndISXVIs";
        break;
    case 390:
        return "BrtBeginColInfos";
        break;
    case 391:
        return "BrtEndColInfos";
        break;
    case 392:
        return "BrtBeginRwBrk";
        break;
    case 393:
        return "BrtEndRwBrk";
        break;
    case 394:
        return "BrtBeginColBrk";
        break;
    case 395:
        return "BrtEndColBrk";
        break;
    case 396:
        return "BrtBrk";
        break;
    case 397:
        return "BrtUserBookView";
        break;
    case 398:
        return "BrtInfo";
        break;
    case 399:
        return "BrtCUsr";
        break;
    case 400:
        return "BrtUsr";
        break;
    case 401:
        return "BrtBeginUsers";
        break;
    case 403:
        return "BrtEOF";
        break;
    case 404:
        return "BrtUCR";
        break;
    case 405:
        return "BrtRRInsDel";
        break;
    case 406:
        return "BrtRREndInsDel";
        break;
    case 407:
        return "BrtRRMove";
        break;
    case 408:
        return "BrtRREndMove";
        break;
    case 409:
        return "BrtRRChgCell";
        break;
    case 410:
        return "BrtRREndChgCell";
        break;
    case 411:
        return "BrtRRHeader";
        break;
    case 412:
        return "BrtRRUserView";
        break;
    case 413:
        return "BrtRRRenSheet";
        break;
    case 414:
        return "BrtRRInsertSh";
        break;
    case 415:
        return "BrtRRDefName";
        break;
    case 416:
        return "BrtRRNote";
        break;
    case 417:
        return "BrtRRConflict";
        break;
    case 418:
        return "BrtRRTQSIF";
        break;
    case 419:
        return "BrtRRFormat";
        break;
    case 420:
        return "BrtRREndFormat";
        break;
    case 421:
        return "BrtRRAutoFmt";
        break;
    case 422:
        return "BrtBeginUserShViews";
        break;
    case 423:
        return "BrtBeginUserShView";
        break;
    case 424:
        return "BrtEndUserShView";
        break;
    case 425:
        return "BrtEndUserShViews";
        break;
    case 426:
        return "BrtArrFmla";
        break;
    case 427:
        return "BrtShrFmla";
        break;
    case 428:
        return "BrtTable";
        break;
    case 429:
        return "BrtBeginExtConnections";
        break;
    case 430:
        return "BrtEndExtConnections";
        break;
    case 431:
        return "BrtBeginPCDCalcMems";
        break;
    case 432:
        return "BrtEndPCDCalcMems";
        break;
    case 433:
        return "BrtBeginPCDCalcMem";
        break;
    case 434:
        return "BrtEndPCDCalcMem";
        break;
    case 435:
        return "BrtBeginPCDHGLevels";
        break;
    case 436:
        return "BrtEndPCDHGLevels";
        break;
    case 437:
        return "BrtBeginPCDHGLevel";
        break;
    case 438:
        return "BrtEndPCDHGLevel";
        break;
    case 439:
        return "BrtBeginPCDHGLGroups";
        break;
    case 440:
        return "BrtEndPCDHGLGroups";
        break;
    case 441:
        return "BrtBeginPCDHGLGroup";
        break;
    case 442:
        return "BrtEndPCDHGLGroup";
        break;
    case 443:
        return "BrtBeginPCDHGLGMembers";
        break;
    case 444:
        return "BrtEndPCDHGLGMembers";
        break;
    case 445:
        return "BrtBeginPCDHGLGMember";
        break;
    case 446:
        return "BrtEndPCDHGLGMember";
        break;
    case 447:
        return "BrtBeginQSI";
        break;
    case 448:
        return "BrtEndQSI";
        break;
    case 449:
        return "BrtBeginQSIR";
        break;
    case 450:
        return "BrtEndQSIR";
        break;
    case 451:
        return "BrtBeginDeletedNames";
        break;
    case 452:
        return "BrtEndDeletedNames";
        break;
    case 453:
        return "BrtBeginDeletedName";
        break;
    case 454:
        return "BrtEndDeletedName";
        break;
    case 455:
        return "BrtBeginQSIFs";
        break;
    case 456:
        return "BrtEndQSIFs";
        break;
    case 457:
        return "BrtBeginQSIF";
        break;
    case 458:
        return "BrtEndQSIF";
        break;
    case 459:
        return "BrtBeginAutoSortScope";
        break;
    case 460:
        return "BrtEndAutoSortScope";
        break;
    case 461:
        return "BrtBeginConditionalFormatting";
        break;
    case 462:
        return "BrtEndConditionalFormatting";
        break;
    case 463:
        return "BrtBeginCFRule";
        break;
    case 464:
        return "BrtEndCFRule";
        break;
    case 465:
        return "BrtBeginIconSet";
        break;
    case 466:
        return "BrtEndIconSet";
        break;
    case 467:
        return "BrtBeginDatabar";
        break;
    case 468:
        return "BrtEndDatabar";
        break;
    case 469:
        return "BrtBeginColorScale";
        break;
    case 470:
        return "BrtEndColorScale";
        break;
    case 471:
        return "BrtCFVO";
        break;
    case 472:
        return "BrtExternValueMeta";
        break;
    case 473:
        return "BrtBeginColorPalette";
        break;
    case 474:
        return "BrtEndColorPalette";
        break;
    case 475:
        return "BrtIndexedColor";
        break;
    case 476:
        return "BrtMargins";
        break;
    case 477:
        return "BrtPrintOptions";
        break;
    case 478:
        return "BrtPageSetup";
        break;
    case 479:
        return "BrtBeginHeaderFooter";
        break;
    case 480:
        return "BrtEndHeaderFooter";
        break;
    case 481:
        return "BrtBeginSXCrtFormat";
        break;
    case 482:
        return "BrtEndSXCrtFormat";
        break;
    case 483:
        return "BrtBeginSXCrtFormats";
        break;
    case 484:
        return "BrtEndSXCrtFormats";
        break;
    case 485:
        return "BrtWsFmtInfo";
        break;
    case 486:
        return "BrtBeginMgs";
        break;
    case 487:
        return "BrtEndMGs";
        break;
    case 488:
        return "BrtBeginMGMaps";
        break;
    case 489:
        return "BrtEndMGMaps";
        break;
    case 490:
        return "BrtBeginMG";
        break;
    case 491:
        return "BrtEndMG";
        break;
    case 492:
        return "BrtBeginMap";
        break;
    case 493:
        return "BrtEndMap";
        break;
    case 494:
        return "BrtHLink";
        break;
    case 495:
        return "BrtBeginDCon";
        break;
    case 496:
        return "BrtEndDCon";
        break;
    case 497:
        return "BrtBeginDRefs";
        break;
    case 498:
        return "BrtEndDRefs";
        break;
    case 499:
        return "BrtDRef";
        break;
    case 500:
        return "BrtBeginScenMan";
        break;
    case 501:
        return "BrtEndScenMan";
        break;
    case 502:
        return "BrtBeginSct";
        break;
    case 503:
        return "BrtEndSct";
        break;
    case 504:
        return "BrtSlc";
        break;
    case 505:
        return "BrtBeginDXFs";
        break;
    case 506:
        return "BrtEndDXFs";
        break;
    case 507:
        return "BrtDXF";
        break;
    case 508:
        return "BrtBeginTableStyles";
        break;
    case 509:
        return "BrtEndTableStyles";
        break;
    case 510:
        return "BrtBeginTableStyle";
        break;
    case 511:
        return "BrtEndTableStyle";
        break;
    case 512:
        return "BrtTableStyleElement";
        break;
    case 513:
        return "BrtTableStyleClient";
        break;
    case 514:
        return "BrtBeginVolDeps";
        break;
    case 515:
        return "BrtEndVolDeps";
        break;
    case 516:
        return "BrtBeginVolType";
        break;
    case 517:
        return "BrtEndVolType";
        break;
    case 518:
        return "BrtBeginVolMain";
        break;
    case 519:
        return "BrtEndVolMain";
        break;
    case 520:
        return "BrtBeginVolTopic";
        break;
    case 521:
        return "BrtEndVolTopic";
        break;
    case 522:
        return "BrtVolSubtopic";
        break;
    case 523:
        return "BrtVolRef";
        break;
    case 524:
        return "BrtVolNum";
        break;
    case 525:
        return "BrtVolErr";
        break;
    case 526:
        return "BrtVolStr";
        break;
    case 527:
        return "BrtVolBool";
        break;
    case 530:
        return "BrtBeginSortState";
        break;
    case 531:
        return "BrtEndSortState";
        break;
    case 532:
        return "BrtBeginSortCond";
        break;
    case 533:
        return "BrtEndSortCond";
        break;
    case 534:
        return "BrtBookProtection";
        break;
    case 535:
        return "BrtSheetProtection";
        break;
    case 536:
        return "BrtRangeProtection";
        break;
    case 537:
        return "BrtPhoneticInfo";
        break;
    case 538:
        return "BrtBeginECTxtWiz";
        break;
    case 539:
        return "BrtEndECTxtWiz";
        break;
    case 540:
        return "BrtBeginECTWFldInfoLst";
        break;
    case 541:
        return "BrtEndECTWFldInfoLst";
        break;
    case 542:
        return "BrtBeginECTwFldInfo";
        break;
    case 548:
        return "BrtFileSharing";
        break;
    case 549:
        return "BrtOleSize";
        break;
    case 550:
        return "BrtDrawing";
        break;
    case 551:
        return "BrtLegacyDrawing";
        break;
    case 552:
        return "BrtLegacyDrawingHF";
        break;
    case 553:
        return "BrtWebOpt";
        break;
    case 554:
        return "BrtBeginWebPubItems";
        break;
    case 555:
        return "BrtEndWebPubItems";
        break;
    case 556:
        return "BrtBeginWebPubItem";
        break;
    case 557:
        return "BrtEndWebPubItem";
        break;
    case 558:
        return "BrtBeginSXCondFmt";
        break;
    case 559:
        return "BrtEndSXCondFmt";
        break;
    case 560:
        return "BrtBeginSXCondFmts";
        break;
    case 561:
        return "BrtEndSXCondFmts";
        break;
    case 562:
        return "BrtBkHim";
        break;
    case 564:
        return "BrtColor";
        break;
    case 565:
        return "BrtBeginIndexedColors";
        break;
    case 566:
        return "BrtEndIndexedColors";
        break;
    case 569:
        return "BrtBeginMRUColors";
        break;
    case 570:
        return "BrtEndMRUColors";
        break;
    case 572:
        return "BrtMRUColor";
        break;
    case 573:
        return "BrtBeginDVals";
        break;
    case 574:
        return "BrtEndDVals";
        break;
    case 577:
        return "BrtSupNameStart";
        break;
    case 578:
        return "BrtSupNameValueStart";
        break;
    case 579:
        return "BrtSupNameValueEnd";
        break;
    case 580:
        return "BrtSupNameNum";
        break;
    case 581:
        return "BrtSupNameErr";
        break;
    case 582:
        return "BrtSupNameSt";
        break;
    case 583:
        return "BrtSupNameNil";
        break;
    case 584:
        return "BrtSupNameBool";
        break;
    case 585:
        return "BrtSupNameFmla";
        break;
    case 586:
        return "BrtSupNameBits";
        break;
    case 587:
        return "BrtSupNameEnd";
        break;
    case 588:
        return "BrtEndSupBook";
        break;
    case 589:
        return "BrtCellSmartTagProperty";
        break;
    case 590:
        return "BrtBeginCellSmartTag";
        break;
    case 591:
        return "BrtEndCellSmartTag";
        break;
    case 592:
        return "BrtBeginCellSmartTags";
        break;
    case 593:
        return "BrtEndCellSmartTags";
        break;
    case 594:
        return "BrtBeginSmartTags";
        break;
    case 595:
        return "BrtEndSmartTags";
        break;
    case 596:
        return "BrtSmartTagType";
        break;
    case 597:
        return "BrtBeginSmartTagTypes";
        break;
    case 598:
        return "BrtEndSmartTagTypes";
        break;
    case 599:
        return "BrtBeginSXFilters";
        break;
    case 600:
        return "BrtEndSXFilters";
        break;
    case 601:
        return "BrtBeginSXFILTER";
        break;
    case 602:
        return "BrtEndSXFilter";
        break;
    case 603:
        return "BrtBeginFills";
        break;
    case 604:
        return "BrtEndFills";
        break;
    case 605:
        return "BrtBeginCellWatches";
        break;
    case 606:
        return "BrtEndCellWatches";
        break;
    case 607:
        return "BrtCellWatch";
        break;
    case 608:
        return "BrtBeginCRErrs";
        break;
    case 609:
        return "BrtEndCRErrs";
        break;
    case 610:
        return "BrtCrashRecErr";
        break;
    case 611:
        return "BrtBeginFonts";
        break;
    case 612:
        return "BrtEndFonts";
        break;
    case 613:
        return "BrtBeginBorders";
        break;
    case 614:
        return "BrtEndBorders";
        break;
    case 615:
        return "BrtBeginFmts";
        break;
    case 616:
        return "BrtEndFmts";
        break;
    case 617:
        return "BrtBeginCellXFs";
        break;
    case 618:
        return "BrtEndCellXFs";
        break;
    case 619:
        return "BrtBeginStyles";
        break;
    case 620:
        return "BrtEndStyles";
        break;
    case 625:
        return "BrtBigName";
        break;
    case 626:
        return "BrtBeginCellStyleXFs";
        break;
    case 627:
        return "BrtEndCellStyleXFs";
        break;
    case 628:
        return "BrtBeginComments";
        break;
    case 629:
        return "BrtEndComments";
        break;
    case 630:
        return "BrtBeginCommentAuthors";
        break;
    case 631:
        return "BrtEndCommentAuthors";
        break;
    case 632:
        return "BrtCommentAuthor";
        break;
    case 633:
        return "BrtBeginCommentList";
        break;
    case 634:
        return "BrtEndCommentList";
        break;
    case 635:
        return "BrtBeginComment";
        break;
    case 636:
        return "BrtEndComment";
        break;
    case 637:
        return "BrtCommentText";
        break;
    case 638:
        return "BrtBeginOleObjects";
        break;
    case 639:
        return "BrtOleObject";
        break;
    case 640:
        return "BrtEndOleObjects";
        break;
    case 641:
        return "BrtBeginSxrules";
        break;
    case 642:
        return "BrtEndSxRules";
        break;
    case 643:
        return "BrtBeginActiveXControls";
        break;
    case 644:
        return "BrtActiveX";
        break;
    case 645:
        return "BrtEndActiveXControls";
        break;
    case 646:
        return "BrtBeginPCDSDTCEMembersSortBy";
        break;
    case 648:
        return "BrtBeginCellIgnoreECs";
        break;
    case 649:
        return "BrtCellIgnoreEC";
        break;
    case 650:
        return "BrtEndCellIgnoreECs";
        break;
    case 651:
        return "BrtCsProp";
        break;
    case 652:
        return "BrtCsPageSetup";
        break;
    case 653:
        return "BrtBeginUserCsViews";
        break;
    case 654:
        return "BrtEndUserCsViews";
        break;
    case 655:
        return "BrtBeginUserCsView";
        break;
    case 656:
        return "BrtEndUserCsView";
        break;
    case 657:
        return "BrtBeginPcdSFCIEntries";
        break;
    case 658:
        return "BrtEndPCDSFCIEntries";
        break;
    case 659:
        return "BrtPCDSFCIEntry";
        break;
    case 660:
        return "BrtBeginListParts";
        break;
    case 661:
        return "BrtListPart";
        break;
    case 662:
        return "BrtEndListParts";
        break;
    case 663:
        return "BrtSheetCalcProp";
        break;
    case 664:
        return "BrtBeginFnGroup";
        break;
    case 665:
        return "BrtFnGroup";
        break;
    case 666:
        return "BrtEndFnGroup";
        break;
    case 667:
        return "BrtSupAddin";
        break;
    case 668:
        return "BrtSXTDMPOrder";
        break;
    case 669:
        return "BrtCsProtection";
        break;
    case 671:
        return "BrtBeginWsSortMap";
        break;
    case 672:
        return "BrtEndWsSortMap";
        break;
    case 673:
        return "BrtBeginRRSort";
        break;
    case 674:
        return "BrtEndRRSort";
        break;
    case 675:
        return "BrtRRSortItem";
        break;
    case 676:
        return "BrtFileSharingIso";
        break;
    case 677:
        return "BrtBookProtectionIso";
        break;
    case 678:
        return "BrtSheetProtectionIso";
        break;
    case 679:
        return "BrtCsProtectionIso";
        break;
    case 680:
        return "BrtRangeProtectionIso";
        break;
    case 681:
        return "BrtDValList";
        break;
    case 1024:
        return "BrtRwDescent";
        break;
    case 1025:
        return "BrtKnownFonts";
        break;
    case 1026:
        return "BrtBeginSXTupleSet";
        break;
    case 1027:
        return "BrtEndSXTupleSet";
        break;
    case 1028:
        return "BrtBeginSXTupleSetHeader";
        break;
    case 1029:
        return "BrtEndSXTupleSetHeader";
        break;
    case 1030:
        return "BrtSXTupleSetHeaderItem";
        break;
    case 1031:
        return "BrtBeginSXTupleSetData";
        break;
    case 1032:
        return "BrtEndSXTupleSetData";
        break;
    case 1033:
        return "BrtBeginSXTupleSetRow";
        break;
    case 1034:
        return "BrtEndSXTupleSetRow";
        break;
    case 1035:
        return "BrtSXTupleSetRowItem";
        break;
    case 1036:
        return "BrtNameExt";
        break;
    case 1037:
        return "BrtPCDH14";
        break;
    case 1038:
        return "BrtBeginPCDCalcMem14";
        break;
    case 1039:
        return "BrtEndPCDCalcMem14";
        break;
    case 1040:
        return "BrtSXTH14";
        break;
    case 1041:
        return "BrtBeginSparklineGroup";
        break;
    case 1042:
        return "BrtEndSparklineGroup";
        break;
    case 1043:
        return "BrtSparkline";
        break;
    case 1044:
        return "BrtSXDI14";
        break;
    case 1045:
        return "BrtWsFmtInfoEx14";
        break;
    case 1046:
        return "BrtBeginConditionalFormatting14";
        break;
    case 1047:
        return "BrtEndConditionalFormatting14";
        break;
    case 1048:
        return "BrtBeginCFRule14";
        break;
    case 1049:
        return "BrtEndCFRule14";
        break;
    case 1050:
        return "BrtCFVO14";
        break;
    case 1051:
        return "BrtBeginDatabar14";
        break;
    case 1052:
        return "BrtBeginIconSet14";
        break;
    case 1053:
        return "BrtDVal14";
        break;
    case 1054:
        return "BrtBeginDVals14";
        break;
    case 1055:
        return "BrtColor14";
        break;
    case 1056:
        return "BrtBeginSparklines";
        break;
    case 1057:
        return "BrtEndSparklines";
        break;
    case 1058:
        return "BrtBeginSparklineGroups";
        break;
    case 1059:
        return "BrtEndSparklineGroups";
        break;
    case 1061:
        return "BrtSXVD14";
        break;
    case 1062:
        return "BrtBeginSxView14";
        break;
    case 1063:
        return "BrtEndSxView14";
        break;
    case 1064:
        return "BrtBeginSXView16";
        break;
    case 1065:
        return "BrtEndSXView16";
        break;
    case 1066:
        return "BrtBeginPCD14";
        break;
    case 1067:
        return "BrtEndPCD14";
        break;
    case 1068:
        return "BrtBeginExtConn14";
        break;
    case 1069:
        return "BrtEndExtConn14";
        break;
    case 1070:
        return "BrtBeginSlicerCacheIDs";
        break;
    case 1071:
        return "BrtEndSlicerCacheIDs";
        break;
    case 1072:
        return "BrtBeginSlicerCacheID";
        break;
    case 1073:
        return "BrtEndSlicerCacheID";
        break;
    case 1075:
        return "BrtBeginSlicerCache";
        break;
    case 1076:
        return "BrtEndSlicerCache";
        break;
    case 1077:
        return "BrtBeginSlicerCacheDef";
        break;
    case 1078:
        return "BrtEndSlicerCacheDef";
        break;
    case 1079:
        return "BrtBeginSlicersEx";
        break;
    case 1080:
        return "BrtEndSlicersEx";
        break;
    case 1081:
        return "BrtBeginSlicerEx";
        break;
    case 1082:
        return "BrtEndSlicerEx";
        break;
    case 1083:
        return "BrtBeginSlicer";
        break;
    case 1084:
        return "BrtEndSlicer";
        break;
    case 1085:
        return "BrtSlicerCachePivotTables";
        break;
    case 1086:
        return "BrtBeginSlicerCacheOlapImpl";
        break;
    case 1087:
        return "BrtEndSlicerCacheOlapImpl";
        break;
    case 1088:
        return "BrtBeginSlicerCacheLevelsData";
        break;
    case 1089:
        return "BrtEndSlicerCacheLevelsData";
        break;
    case 1090:
        return "BrtBeginSlicerCacheLevelData";
        break;
    case 1091:
        return "BrtEndSlicerCacheLevelData";
        break;
    case 1092:
        return "BrtBeginSlicerCacheSiRanges";
        break;
    case 1093:
        return "BrtEndSlicerCacheSiRanges";
        break;
    case 1094:
        return "BrtBeginSlicerCacheSiRange";
        break;
    case 1095:
        return "BrtEndSlicerCacheSiRange";
        break;
    case 1096:
        return "BrtSlicerCacheOlapItem";
        break;
    case 1097:
        return "BrtBeginSlicerCacheSelections";
        break;
    case 1098:
        return "BrtSlicerCacheSelection";
        break;
    case 1099:
        return "BrtEndSlicerCacheSelections";
        break;
    case 1100:
        return "BrtBeginSlicerCacheNative";
        break;
    case 1101:
        return "BrtEndSlicerCacheNative";
        break;
    case 1102:
        return "BrtSlicerCacheNativeItem";
        break;
    case 1103:
        return "BrtRangeProtection14";
        break;
    case 1104:
        return "BrtRangeProtectionIso14";
        break;
    case 1105:
        return "BrtCellIgnoreEC14";
        break;
    case 1111:
        return "BrtList14";
        break;
    case 1112:
        return "BrtCFIcon";
        break;
    case 1113:
        return "BrtBeginSlicerCachesPivotCacheIDs";
        break;
    case 1114:
        return "BrtEndSlicerCachesPivotCacheIDs";
        break;
    case 1115:
        return "BrtBeginSlicers";
        break;
    case 1116:
        return "BrtEndSlicers";
        break;
    case 1117:
        return "BrtWbProp14";
        break;
    case 1118:
        return "BrtBeginSXEdit";
        break;
    case 1119:
        return "BrtEndSXEdit";
        break;
    case 1120:
        return "BrtBeginSXEdits";
        break;
    case 1121:
        return "BrtEndSXEdits";
        break;
    case 1122:
        return "BrtBeginSXChange";
        break;
    case 1123:
        return "BrtEndSXChange";
        break;
    case 1124:
        return "BrtBeginSXChanges";
        break;
    case 1125:
        return "BrtEndSXChanges";
        break;
    case 1126:
        return "BrtSXTupleItems";
        break;
    case 1128:
        return "BrtBeginSlicerStyle";
        break;
    case 1129:
        return "BrtEndSlicerStyle";
        break;
    case 1130:
        return "BrtSlicerStyleElement";
        break;
    case 1131:
        return "BrtBeginStyleSheetExt14";
        break;
    case 1132:
        return "BrtEndStyleSheetExt14";
        break;
    case 1133:
        return "BrtBeginSlicerCachesPivotCacheID";
        break;
    case 1134:
        return "BrtEndSlicerCachesPivotCacheID";
        break;
    case 1135:
        return "BrtBeginConditionalFormattings";
        break;
    case 1136:
        return "BrtEndConditionalFormattings";
        break;
    case 1137:
        return "BrtBeginPCDCalcMemExt";
        break;
    case 1138:
        return "BrtEndPCDCalcMemExt";
        break;
    case 1139:
        return "BrtBeginPCDCalcMemsExt";
        break;
    case 1140:
        return "BrtEndPCDCalcMemsExt";
        break;
    case 1141:
        return "BrtPCDField14";
        break;
    case 1142:
        return "BrtBeginSlicerStyles";
        break;
    case 1143:
        return "BrtEndSlicerStyles";
        break;
    case 1144:
        return "BrtBeginSlicerStyleElements";
        break;
    case 1145:
        return "BrtEndSlicerStyleElements";
        break;
    case 1146:
        return "BrtCFRuleExt";
        break;
    case 1147:
        return "BrtBeginSXCondFmt14";
        break;
    case 1148:
        return "BrtEndSXCondFmt14";
        break;
    case 1149:
        return "BrtBeginSXCondFmts14";
        break;
    case 1150:
        return "BrtEndSXCondFmts14";
        break;
    case 1152:
        return "BrtBeginSortCond14";
        break;
    case 1153:
        return "BrtEndSortCond14";
        break;
    case 1154:
        return "BrtEndDVals14";
        break;
    case 1155:
        return "BrtEndIconSet14";
        break;
    case 1156:
        return "BrtEndDatabar14";
        break;
    case 1157:
        return "BrtBeginColorScale14";
        break;
    case 1158:
        return "BrtEndColorScale14";
        break;
    case 1159:
        return "BrtBeginSxrules14";
        break;
    case 1160:
        return "BrtEndSxrules14";
        break;
    case 1161:
        return "BrtBeginPRule14";
        break;
    case 1162:
        return "BrtEndPRule14";
        break;
    case 1163:
        return "BrtBeginPRFilters14";
        break;
    case 1164:
        return "BrtEndPRFilters14";
        break;
    case 1165:
        return "BrtBeginPRFilter14";
        break;
    case 1166:
        return "BrtEndPRFilter14";
        break;
    case 1167:
        return "BrtBeginPRFItem14";
        break;
    case 1168:
        return "BrtEndPRFItem14";
        break;
    case 1169:
        return "BrtBeginCellIgnoreECs14";
        break;
    case 1170:
        return "BrtEndCellIgnoreECs14";
        break;
    case 1171:
        return "BrtDxf14";
        break;
    case 1172:
        return "BrtBeginDxF14s";
        break;
    case 1173:
        return "BrtEndDxf14s";
        break;
    case 1177:
        return "BrtFilter14";
        break;
    case 1178:
        return "BrtBeginCustomFilters14";
        break;
    case 1180:
        return "BrtCustomFilter14";
        break;
    case 1181:
        return "BrtIconFilter14";
        break;
    case 1182:
        return "BrtPivotCacheConnectionName";
        break;
    case 2048:
        return "BrtBeginDecoupledPivotCacheIDs";
        break;
    case 2049:
        return "BrtEndDecoupledPivotCacheIDs";
        break;
    case 2050:
        return "BrtDecoupledPivotCacheID";
        break;
    case 2051:
        return "BrtBeginPivotTableRefs";
        break;
    case 2052:
        return "BrtEndPivotTableRefs";
        break;
    case 2053:
        return "BrtPivotTableRef";
        break;
    case 2054:
        return "BrtSlicerCacheBookPivotTables";
        break;
    case 2055:
        return "BrtBeginSxvcells";
        break;
    case 2056:
        return "BrtEndSxvcells";
        break;
    case 2057:
        return "BrtBeginSxRow";
        break;
    case 2058:
        return "BrtEndSxRow";
        break;
    case 2060:
        return "BrtPcdCalcMem15";
        break;
    case 2067:
        return "BrtQsi15";
        break;
    case 2068:
        return "BrtBeginWebExtensions";
        break;
    case 2069:
        return "BrtEndWebExtensions";
        break;
    case 2070:
        return "BrtWebExtension";
        break;
    case 2071:
        return "BrtAbsPath15";
        break;
    case 2072:
        return "BrtBeginPivotTableUISettings";
        break;
    case 2073:
        return "BrtEndPivotTableUISettings";
        break;
    case 2075:
        return "BrtTableSlicerCacheIDs";
        break;
    case 2076:
        return "BrtTableSlicerCacheID";
        break;
    case 2077:
        return "BrtBeginTableSlicerCache";
        break;
    case 2078:
        return "BrtEndTableSlicerCache";
        break;
    case 2079:
        return "BrtSxFilter15";
        break;
    case 2080:
        return "BrtBeginTimelineCachePivotCacheIDs";
        break;
    case 2081:
        return "BrtEndTimelineCachePivotCacheIDs";
        break;
    case 2082:
        return "BrtTimelineCachePivotCacheID";
        break;
    case 2083:
        return "BrtBeginTimelineCacheIDs";
        break;
    case 2084:
        return "BrtEndTimelineCacheIDs";
        break;
    case 2085:
        return "BrtBeginTimelineCacheID";
        break;
    case 2086:
        return "BrtEndTimelineCacheID";
        break;
    case 2087:
        return "BrtBeginTimelinesEx";
        break;
    case 2088:
        return "BrtEndTimelinesEx";
        break;
    case 2089:
        return "BrtBeginTimelineEx";
        break;
    case 2090:
        return "BrtEndTimelineEx";
        break;
    case 2091:
        return "BrtWorkBookPr15";
        break;
    case 2092:
        return "BrtPCDH15";
        break;
    case 2093:
        return "BrtBeginTimelineStyle";
        break;
    case 2094:
        return "BrtEndTimelineStyle";
        break;
    case 2095:
        return "BrtTimelineStyleElement";
        break;
    case 2096:
        return "BrtBeginTimelineStylesheetExt15";
        break;
    case 2097:
        return "BrtEndTimelineStylesheetExt15";
        break;
    case 2098:
        return "BrtBeginTimelineStyles";
        break;
    case 2099:
        return "BrtEndTimelineStyles";
        break;
    case 2100:
        return "BrtBeginTimelineStyleElements";
        break;
    case 2101:
        return "BrtEndTimelineStyleElements";
        break;
    case 2102:
        return "BrtDxf15";
        break;
    case 2103:
        return "BrtBeginDxfs15";
        break;
    case 2104:
        return "BrtEndDXFs15";
        break;
    case 2105:
        return "BrtSlicerCacheHideItemsWithNoData";
        break;
    case 2106:
        return "BrtBeginItemUniqueNames";
        break;
    case 2107:
        return "BrtEndItemUniqueNames";
        break;
    case 2108:
        return "BrtItemUniqueName";
        break;
    case 2109:
        return "BrtBeginExtConn15";
        break;
    case 2110:
        return "BrtEndExtConn15";
        break;
    case 2111:
        return "BrtBeginOledbPr15";
        break;
    case 2112:
        return "BrtEndOledbPr15";
        break;
    case 2113:
        return "BrtBeginDataFeedPr15";
        break;
    case 2114:
        return "BrtEndDataFeedPr15";
        break;
    case 2115:
        return "BrtTextPr15";
        break;
    case 2116:
        return "BrtRangePr15";
        break;
    case 2117:
        return "BrtDbCommand15";
        break;
    case 2118:
        return "BrtBeginDbTables15";
        break;
    case 2119:
        return "BrtEndDbTables15";
        break;
    case 2120:
        return "BrtDbTable15";
        break;
    case 2121:
        return "BrtBeginDataModel";
        break;
    case 2122:
        return "BrtEndDataModel";
        break;
    case 2123:
        return "BrtBeginModelTables";
        break;
    case 2124:
        return "BrtEndModelTables";
        break;
    case 2125:
        return "BrtModelTable";
        break;
    case 2126:
        return "BrtBeginModelRelationships";
        break;
    case 2127:
        return "BrtEndModelRelationships";
        break;
    case 2128:
        return "BrtModelRelationship";
        break;
    case 2129:
        return "BrtBeginECTxtWiz15";
        break;
    case 2130:
        return "BrtEndECTxtWiz15";
        break;
    case 2131:
        return "BrtBeginECTWFldInfoLst15";
        break;
    case 2132:
        return "BrtEndECTWFldInfoLst15";
        break;
    case 2133:
        return "BrtBeginECTWFldInfo15";
        break;
    case 2134:
        return "BrtFieldListActiveItem";
        break;
    case 2135:
        return "BrtPivotCacheIdVersion";
        break;
    case 2136:
        return "BrtSXDI15";
        break;
    case 2137:
        return "brtBeginModelTimeGroupings";
        break;
    case 2138:
        return "brtEndModelTimeGroupings";
        break;
    case 2139:
        return "brtBeginModelTimeGrouping";
        break;
    case 2140:
        return "brtEndModelTimeGrouping";
        break;
    case 2141:
        return "brtModelTimeGroupingCalcCol";
        break;
    case 3073:
        return "brtRevisionPtr";
        break;
    case 4096:
        return "BrtBeginDynamicArrayPr";
        break;
    case 4097:
        return "BrtEndDynamicArrayPr";
        break;
    case 5002:
        return "BrtBeginRichValueBlock";
        break;
    case 5003:
        return "BrtEndRichValueBlock";
        break;
    case 5081:
        return "BrtBeginRichFilters";
        break;
    case 5082:
        return "BrtEndRichFilters";
        break;
    case 5083:
        return "BrtRichFilter";
        break;
    case 5084:
        return "BrtBeginRichFilterColumn";
        break;
    case 5085:
        return "BrtEndRichFilterColumn";
        break;
    case 5086:
        return "BrtBeginCustomRichFilters";
        break;
    case 5087:
        return "BrtEndCustomRichFilters";
        break;
    case 5088:
        return "BRTCustomRichFilter";
        break;
    case 5089:
        return "BrtTop10RichFilter";
        break;
    case 5090:
        return "BrtDynamicRichFilter";
        break;
    case 5092:
        return "BrtBeginRichSortCondition";
        break;
    case 5093:
        return "BrtEndRichSortCondition";
        break;
    case 5094:
        return "BrtRichFilterDateGroupItem";
        break;
    case 5095:
        return "BrtBeginCalcFeatures";
        break;
    case 5096:
        return "BrtEndCalcFeatures";
        break;
    case 5097:
        return "BrtCalcFeature";
        break;
    default:
        return "Unknown";
    }
}

} // namespace readxlsb