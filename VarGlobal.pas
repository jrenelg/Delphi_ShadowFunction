unit VarGlobal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, FileCtrl, Gauges, StrUtils, ADODB,
  ExtCtrls, Inifiles, Outline, shellapi, ShlObj, ActiveX,
  DB, DBTables, CheckLst, ComCtrls, Buttons, Grids, IdFTP;

// procedure ShF_CleanReg;
procedure ShF_Inicio;
procedure ShF_Fin;
procedure ShF_TablasDeIni();
procedure ShF_CheckBitac();
function ShF_SQL(IDSQL: Integer; Var001: String = ''; Var002: String = '';
  Var003: String = ''): String;
function ShF_ZipBienCreado(FileZIP: String): boolean;

var

  glbRutaEXE, glbDirRaizSys, glbNameFileError, glbError, glbSessionName: String;
  glbContError: Integer;
  glbDirConfig, glbDirError, glbDirTMP, glbFileTAUX, glbINIConfig: string;
  glbBanSession: boolean;

  // Control de DB
  glbTSession: TSession;
  glbQryBDE01, glbQryBDE02, glbQryBDE03: TQuery;
  glbTTable01, glbTTable02: TTable;
  glbTBachMove01: TBatchMove;
  glbTADOTable01: TADOTable;
  glbQryADO01, glbQryADO02: TADOQuery;
  glbCONN: TADOConnection;
  glbSQL: WideString;

  // Variable General
  glbCont: Integer;
  glbIniFile: TIniFile;
  glbDate: TDateTime;
  glbDateStr: string;
  glbDateFloat: Double;
  // Variables Sistema CAPI
  glbDirEstruct, glbDirSysActualiza, glbDirSys, // carpeta sistema
  glbDirFileSys, // carpeta archivos de sistema
  glbDirUsu, // carpeta usuarios
  glbDirAutoguardado, // carpeta Autoguardado
  glbDirTemporal, glbINIEnc, glbINIVer, glbExeAdmin, // archivo ini
  glbUserActivo, glbUserARecibir: string;

  glbTablaUsu, glbAuxTablaUsu, glbTablaSys: TStringList;
  // lista de archivos de usuario y sistema (dbfs)

  glbTabMov, // tabla Movimientos
  glbTabListMov, // tabla lista movimientos
  glbTabUsuario, // tabla usuarios
  glbTabVivienda, // tabla vivienda
  glbTabResEntrevista // tabala resentrevista
    : string;

  // variables que componene clave de usuario y tipo
  glbOficina, glbAUX, glbJE, glbENT, glbTipoUsu: string;

  glbControl, glbViv_Sel, glbCve_Usu, glbNum_Vis: String;
  glbRaizMem: String;

  glbListaZip, glbLlave: TStringList;

const

  // Constantes Sistema
  glbNameSys = 'Nombre Proyecto';
  glbNameEnc = 'Proyecto';
  glbNameEXE = 'Proyecto.exe';
  glbMostrarErrores = True;
  glbWebINEGI = 'http://www.inegi.org.mx';

  glbConst_DirTemporal = '\Temporal';
  glbConst_DirConfig = '\FilesConfig';
  glbConst_DirError = glbConst_DirConfig + '\HistorialDeError';
  glbConst_DirTMP = glbConst_DirConfig + glbConst_DirTemporal;
  glbConst_FileTAUX = glbConst_DirConfig + '\BASE.DBF';
  glbConst_FileIniConf = glbConst_DirConfig + '\Config.ini';

  // Dir y Archivos Sistema CAPI


  // Estructura Dir

  glbZipORG = '.zip';
  glbZipEXT = '.ctr';

  // FTP externo
  glbFTPUser = 'enifext';
  glbFTPPass = 'QUrOH3V';
  glbFTPHost = 'ftpexterno.inegi.gob.mx';
  glbFTPHostSubDir = 'enifext';

implementation

Uses ShadowFunction;

procedure ShF_Inicio;
var
  Sem: THandle;
begin
  Sem := CreateSemaphore(nil, 0, 1, glbNameEXE);
  if ((Sem <> 0) and (GetLastError = ERROR_ALREADY_EXISTS)) then
  begin
    CloseHandle(Sem);
    Halt;
    Exit;
  end;
  ShF_MaxPriori;
  glbBanSession := False;
  glbTSession := TSession.Create(nil);
  glbSessionName := ShF_Session_ON(glbTSession, glbNameSys + '_');
  glbRutaEXE := GetCurrentDir;
  glbDirRaizSys := ShF_DirUp(glbDirRaizSys);
  glbDirConfig := glbDirRaizSys + glbConst_DirConfig;
  glbDirError := glbDirRaizSys + glbConst_DirError;
  glbDirTMP := glbDirRaizSys + glbConst_DirTMP;
  glbFileTAUX := glbDirRaizSys + glbConst_FileTAUX;
  glbINIConfig := glbDirRaizSys + glbConst_FileIniConf;
  glbNameFileError := glbDirError + '\' + ShF_GenerarNombreConFecha
    ('CSA') + '.bit';
  glbContError := 0;

end;

procedure ShF_Fin;
begin
  ShF_Session_OFF(glbTSession, glbSessionName);
  glbBanSession := False;
end;

// Procedimiento que saca de dos cadenas de caracteres de rubros de un INI las tablas del sistema
procedure ShF_TablasDeIni();
begin
  try
    glbTablaUsu.Clear;
    glbAuxTablaUsu.Clear;
    glbTablaSys.Clear;
    glbTablaUsu := ShF_CAPI_ListaINI('Tabla_Usu');
    glbAuxTablaUsu := ShF_CAPI_ListaINI('TablaAux_Usu');
    glbTablaSys := ShF_CAPI_ListaINI('Tabla_Sys');
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_TablasDeIni][' + E.Message + ']');
  end;
end;

procedure ShF_MigrarRegistro(Origen, Destino, LLave: String);
var
  QryA: TQuery;
begin
  try
    try
      QryA := TQuery.Create(nil);
      if ShF_BDE_OpenSelect(QryA, '*', Origen, LLave) then
      begin
        QryA.First;
        ShF_BDE_Insert_All(Destino, QryA);
      end;
      ShF_BDE_CloseSQL(QryA);
    except
      ShowMessage('[E.ShF.ControlSQL] ShF_MigrarRegistro');
    end;
  finally

  end;
end;

function ShF_ZipBienCreado(FileZIP: String): boolean;
var
  Bien: boolean;
  F: file of byte;
begin
  Bien := True;
  AssignFile(F, FileZIP);
  Reset(F);
  if FileSize(F) <= 0 then
    Bien := False;
  CloseFile(F);
  Result := Bien;
end;

procedure ShF_CheckBitac();
begin
  if not FileExists(glbTabMov) then
  begin
    glbSQL := 'SELECT ';
    glbSQL := glbSQL + ' N10 as N_MOV,';
    glbSQL := glbSQL + ' C12 as FECHA,';
    glbSQL := glbSQL + ' C20 as FECHAF,';
    glbSQL := glbSQL + ' C5 as TIPO,';
    glbSQL := glbSQL + ' C10 as ORIGEN,';
    glbSQL := glbSQL + ' C10 as DESTINO,';
    glbSQL := glbSQL + ' C50 as NOMFILE ';
    glbSQL := glbSQL + ' FROM "' + glbFileTAUX + '"';
    glbQryBDE01 := TQuery.Create(nil);
    if not ShF_BDE_OpenSQL(glbSQL, glbQryBDE01) then
    begin
      Shf_CrearTabla(glbTabMov, glbQryBDE01, ttDBase);
    end;
    ShF_BDE_CloseSQL(glbQryBDE01);
  end;
  if not FileExists(glbTabListMov) then
  begin
    glbSQL := 'SELECT ';
    glbSQL := glbSQL + ' N10 as N_MOV,';
    glbSQL := glbSQL + ' C15 as CONTROL,';
    glbSQL := glbSQL + ' C5 as VIV_SEL,';
    glbSQL := glbSQL + ' C5 as RESUL_V,';
    glbSQL := glbSQL + ' C10 as ACC,';
    glbSQL := glbSQL + ' C10 as DESTINO ';
    glbSQL := glbSQL + ' FROM "' + glbFileTAUX + '"';
    glbQryBDE01 := TQuery.Create(nil);
    if not ShF_BDE_OpenSQL(glbSQL, glbQryBDE01) then
    begin
      Shf_CrearTabla(glbTabListMov, glbQryBDE01, ttDBase);
    end;
    ShF_BDE_CloseSQL(glbQryBDE01);
  end;
end;

function ShF_SQL(IDSQL: Integer; Var001: String = ''; Var002: String = '';
  Var003: String = ''): String;
var
  SQL: String;
begin
  try
    try
      SQL := '';
      case IDSQL of
        // Seleccion de todos los Usuarios de esa jefatura sin contar al activo
        1:
          SQL := 'SELECT cve_usu FROM "' + glbTabUsuario +
            '" WHERE cve_usu LIKE "' + glbOficina + glbAUX + glbJE +
            '%" AND cve_usu <> "' + glbUserActivo + '"';
        // Selecciona todos los usuarios de esa entidad
        2:
          SQL := 'SELECT cve_usu FROM "' + glbTabUsuario +
            '" WHERE cve_usu LIKE "' + glbOficina + '%"';
        // Seleccionar todas la viviendas en la carpeta de ese usuario
        3:
          SQL := 'SELECT Control, Viv_Sel, Resul_V FROM "' + Var001 + '\' +
            glbTabVivienda +
            '" GROUP BY Control, Viv_Sel, Resul_V ORDER BY Control, Viv_Sel, Resul_V';
        // Obtener vuelta de usuario
        4:
          SQL := 'SELECT Max(Cast(Vuelta as Int)) as ID FROM "' + glbTabUsuario
            + '" WHERE Cve_usu="' + glbUserARecibir + '"';
        // Obtener usuario con like
        5:
          SQL := 'SELECT cve_usu FROM "' + glbTabUsuario +
            '" WHERE cve_usu LIKE "' + glbUserARecibir + '%"';
        // Obtener usuario con like
        6:
          SQL := 'SELECT cve_usu, UACT FROM "' + glbTabUsuario +
            '" WHERE cve_usu LIKE "' + glbOficina + '%0" AND not cve_usu LIKE "'
            + glbOficina + '%00"';
        //
        7:
          SQL := 'SELECT distinct origen FROM "' + glbTabMov + '"';
        //
        8:
          begin
            SQL := 'select origen, destino, control, viv_sel, resul_v, acc';
            SQL := SQL + ' from "' + glbTabMov + '" as M';
            SQL := SQL + ' inner join "' + glbTabListMov + '" as L';
            SQL := SQL + ' on M.n_mov = L.n_mov ';
            if Length(Trim(Var001)) > 0 then
              SQL := SQL + ' where ' + Var001;
            SQL := SQL + ' Order by  control, viv_sel';
          end;
        //
        9:
          SQL := 'SELECT distinct Cast(Fecha as date) as Fecha FROM "' +
            glbTabMov + '" WHERE Tipo = "' + Var001 + '" order by fecha desc';
        10:
          SQL := 'SELECT distinct Hra FROM "' + glbTabMov + '" WHERE Tipo = "' +
            Var001 + '" AND Fecha = "' + Var002 + '" order by hra desc';
        11:
          SQL := 'SELECT Max(Cast(n_mov as Int)) as ID FROM "' +
            glbTabMov + '"';
        12:
          SQL := 'SELECT num_vis , fecha , Hra_2 , Resul_V FROM "' + Var001 +
            '" WHERE Control = ''' + Var002 + ''' AND Viv_Sel = ''' + Var003 +
            ''' AND RESUL_V <> ""  Order By num_vis desc, fecha desc, Hra_2 desc, Resul_V';
        13:
          SQL := 'SELECT Vuelta FROM "' + glbTabUsuario + '" WHERE Cve_usu = "'
            + Var001 + '" AND vuelta <> ""';
        14:
          SQL := 'SELECT id_proceso FROM "' + Var001 + '\' + glbTabVivienda +
            '" WHERE id_proceso IN ("RC", "LVCA", "LVCM")';
        15:
          begin
            SQL := 'SELECT TV.Control, TV.Viv_Sel, TV.id_proceso, TV.con_rc, TV.obs_jefe, UV.num_vis, UV.fecha, UV.hra_2,  UV.resul_v';
            SQL := SQL + ' FROM "' + Var001 + '\' + glbTabVivienda + '" As TV';
            SQL := SQL + ' INNER JOIN "' + Var002 + '" As UV';
            SQL := SQL +
              ' ON TV.Control = UV.Control AND TV.Viv_Sel = UV.Viv_Sel';
            SQL := SQL + ' WHERE TV.id_proceso IN ("RC", "LVCA", "LVCM")';
          end;
        16:
          begin
            SQL := 'SELECT RV.Control, RV.Viv_Sel, RV.id_proceso, RV.con_rc, RV.obs_jefe, RV.num_vis As num_vis_rv, RV.fecha As fecha_rv, RV.hra_2 As hra_2_rv, RV.Resul_V , UV.num_vis As num_vis_uv , UV.fecha As fecha_uv, UV.hra_2 As hra_2_uv';
            SQL := SQL + ' FROM "' + Var001 + '\Resumen_RV.dbf" As RV';
            SQL := SQL + ' INNER JOIN "' + Var002 + '" As UV';
            SQL := SQL +
              ' ON RV.Control = UV.Control AND RV.Viv_Sel = UV.Viv_Sel';
            SQL := SQL + ' WHERE RV.id_proceso IN ("RC", "LVCA", "LVCM")';
          end;
        17:
          SQL := 'SELECT UACT FROM "' + glbTabUsuario + '" WHERE Cve_usu = "' +
            Var001 + '"';
        18:
          SQL := 'SELECT cve_usu, UACT FROM "' + glbTabUsuario +
            '" WHERE cve_usu LIKE "' + glbUserARecibir + '%"';
        19:
          begin
            SQL := 'select distinct T1.Control, T1.Viv_Sel, T1.Resul_V';
            SQL := SQL + ' from "' + Var001 + '\' + glbTabVivienda + '" as T1';
            SQL := SQL + ' LEFT JOIN "' + Var002 + '\' + glbTabVivienda +
              '" as T2';
            SQL := SQL + ' ON T1.Control=T2.Control and T1.Viv_Sel =T2.Viv_Sel';
            // if not ForzarSinc then
            SQL := SQL +
              ' Where (Cast(T1.Num_Vis as Int) > Cast(T2.Num_Vis as int)) OR (Cast(T1.Num_Vis as Int) >0 and (T2.Num_vis is NULL OR Trim(T2.Num_vis) =""))';
          end;
        20:
          begin
            SQL := 'select T.* From "' + Var001 + '\' + Var002 + '" as T';
            SQL := SQL + ' Where T.Control+T.Viv_Sel IN ';
            SQL := SQL + ' (select distinct T1.Control+T1.Viv_Sel';
            SQL := SQL + ' from "' + Var001 + '\Viv_Act.dbf" as T1)';
          end;
        21:
          begin
            SQL := 'Delete From "' + Var002 + '\' + Var003 + '" as T';
            SQL := SQL + ' Where T.Control+T.Viv_Sel IN ';
            SQL := SQL + ' (select distinct T1.Control+T1.Viv_Sel';
            SQL := SQL + ' from "' + Var001 + '\Viv_Act.dbf" as T1)';
          end;
        22:
          SQL := 'SELECT V0_6 FROM "' + glbTabUsuario + '" WHERE Cve_usu = "' +
            Var001 + '"';
      end;
    except
      on E: Exception do
        ShF_ControlDeError('[VarGlobal.ShF_SQL][' + E.Message + '] [IDSQL ' +
          IntToStr(IDSQL) + ']');
    end;
  finally
    Result := SQL;
  end;
end;

end.
