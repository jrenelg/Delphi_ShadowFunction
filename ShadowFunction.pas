unit ShadowFunction;

interface

uses
  DB, Math, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
  Forms, Dialogs, ExcelXP, registry,
  OleServer, Buttons, Grids, StdCtrls, DBTables, StrUtils, ADODB, FileCtrl,
  ShellApi, Mask, ExtCtrls, DateUtils,
  ComCtrls, IdFTP, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  inifiles, VarGlobal, ComObj, BDE, AbBase, AbBrowse, AbZBrows, AbZipper,
  AbZipKit,
  AbArcTyp, AbUtils, AbZipOut, AbMeter, AbConst, AbDlgDir, AbDlgPwd, AbZipTyp,
  Outline, AbUnzper, TlHelp32, JvRichEdit, WinInet;

{ ************************************************
  ************** CONTROL SQL **********************
  ************************************************* }
// Pack a una tabla
procedure ShF_Pack(RutTabla: String);
// Abre Conexion y ejecuta query con componente BDE (Selects)
Function ShF_BDE_OpenSQL(SQL: String; Qry: TQuery): Boolean;
// Cierra Conexion con componente BDE
procedure ShF_BDE_CloseSQL(Qry: TQuery);
// Abre Conexion y Ejecuta Sentencia con componentes BDE (Updates, Deletes, Alter)
procedure ShF_BDE_EjecutarSQL(SQL: String);
// Abre Conexion y ejecuta query con componente BDE segun las condiciones indicadas
Function ShF_BDE_OpenSelect(TQry: TQuery; Select: String; From: string;
  Where: String = ''; Condicion: String = ''): Boolean;
// Inserta Un TQry en una tabla con la misma estructura
procedure ShF_BDE_Insert_First(Tabla: string; Qry: TQuery);
procedure ShF_BDE_Insert_All(Tabla: string; Qry: TQuery);
// Establece el directorio del TADOConnection y abre la conexion (Tipo 1 MSDASQL Tipo 2 VFPOLEDB)
procedure ShF_ADO_OpenCONN(Dir: String; ADO: TADOConnection; Tipo: Integer);
// Establece la conexion con el servidor SQL proporcionando instancia y nombre de la DB
procedure ShF_ADO_OpenCONN_SQLSERV(Serv: String; DB: String;
  ADO: TADOConnection; ConAuten: Boolean = False;
  User: String = ''; Pass: String = '');
// Cierra conexion del TADOConnection
procedure ShF_ADO_CloseCONN(ADO: TADOConnection);
// Abre Conexion y ejecuta query con componente ADO
Function ShF_ADO_OpenSQL(SQL: String; Qry: TADOQuery): Boolean;
// Cierra Conexion con componente ADO
procedure ShF_ADO_CloseSQL(Qry: TADOQuery);
// Ejecuta sentencia con componente ADO
procedure ShF_ADO_EjecutarSQL(SQL: String; Serv: String; DB: String;
  ConAuten: Boolean = False; User: String = ''; Pass: String = '');
// Abre Conexion con SQL Server y ejecuta query con componente ADO segun las condiciones indicadas
Function ShF_ADO_OpenSelect(TQry: TADOQuery; Select: String; From: string;
  Where: String = ''; Condicion: String = ''): Boolean;
// Inserta una ADO Qry a una tabla con la misma estructura
procedure ShF_Insert_ADO(Tabla: string; Qry: TADOQuery; Serv: String;
  DB: String);
// Regresa True si todos los campos de un registro (TQuery) estan vacios
Function ShF_RowRSVacio(Qry: TQuery): Boolean;
// Numero de Campos con informacion
Function ShF_CamposNoVacios(Qry: TQuery): Integer;
// Migra todos los registros segun una llave (Folio) a una tabla limpia
procedure ShF_MigrarTablaAllCamposString(oldTab: String; newTab: String;
  key: String);
// de yn TQry Crea un where con todos los campos del tqry
function ShF_Where(Qry: TQuery): String;
// Borra todos los registros de una tabla y los remplaza con los de otra basado en una llave
procedure ShF_ActualziarRegistros(Origen, Destino, LLave: String);
// Crea una tabla en base a la estructura de un TQry
procedure ShF_CrearTabla(NameTabla: string; Qry: TQuery; Tipo: TTableType);
// Pasa un Qry a una tabla DBF
procedure ShF_TQryToDBF(NameTabla: string; Qry: TQuery; Tipo: TTableType);
// Cuantos registros hay en una tabla
function ShF_CuantosRegistros(Tablas: String): String;

{ ************************************************
  **************** UTILERIAS **********************
  ************************************************* }
// Guarda una linea en un TXT
procedure ShF_FileError(Nombre, Error: String);
// Control de errores
procedure ShF_ControlDeError(Error: String);
// Libera memoria
procedure ShF_LiberarMemoria;
// Regresar un dir
Function ShF_DirUp(Ruta: String): String;
// Cuanta cuantos caracteres hay en una cadena
Function ShF_CuantosCharHay(Cadena: String; Caracter: Char): Integer;
// obtener una subcadena o serie de subcadenas de una cadena.
// La función recibe una cadena formateada para tal propósito,
// la subcadena en un índice determinado y opcionalmente un separador
function ShF_GetPosiStrig(s: string; Indice: Byte; Separador: string): string;
// Borrar una carpeta
function ShF_BorrarCarpeta(const rutaCarpeta: string): Boolean;
// ????????
// Elimina los campos no permitidos en un String
Function ShF_LimpiarString(Campo: String; Acentos: Boolean): String;
// Lista Directoros
function ShF_ListaDirectorios(directorioPadre: string;
  ComenzarCon: string = ''): TStringList;
// Lista Archivos
function ShF_ListaArchivos(directorioPadre: string): TStringList;
// Limpa un StringGrid}
Procedure ShF_InicializarStringGrid(SGrid: TStringGrid);
// Descomprime Archivo ZIP
Function ShF_UnZipFile(DirArchivo: string; DirDestino: String): Boolean;
// Copia todo de una ruta a otra
Function ShF_CopiaTodo(Origen, Destino: String): LongInt;
// Lee Archivo INI
Function ShF_LeerINI(Dir, Secc, Vari: string): String;
// Escribe en Archivo INI
Function ShF_EscribirINI(Dir, Secc, Vari, Valor: string): Boolean;
// Cantidad en 0000.00 formato y 1 pesos convierte a letra
Function ShF_CantidadEnLetra(curCantidad: Currency; MonNal: Integer): String;
// devuelven true o false dependiendo de si el texto que buscamos
// se ha encontrado o no. Nos da la opción de limpiar la cadena (strip)
// o hacer o no diferencia entre mayúsculas y minúsculas.
function ShF_FindInStr(Text, Destination: String;
  CaseSensitive: Boolean): Boolean;
function ShF_FindInTextFile(Text, Filepath: String;
  CaseSensitive: Boolean): Boolean;
// limpiar strings de caracteres especiales
function ShF_Strip(s: string; tildes: Boolean): string;
// De un directorio lista todos los archivoz Zip en un stringlist
function ShF_ListaArchivosZip(directorioPadre: string): TStringList;
// Hora en formato hh:mm:ss a.m./p.m.
function ShF_12HrPunto(): String;
// Fecha en Formato dd/mm/aa y hora en hh:mm:ss am/pm
function ShF_Fecha_12HrsSinPunto(Tipo: String): String;
// De una Fecha dd/mm/aaaa saca la edad de la persona
function ShF_Edad(FechaNacimiento: string): Integer;
// Dias Transcurridos entre dos fechas
function ShF_DiasTranscurridos(FechaIni, FechaFin: string): Integer;
// A una Fecha dada sumar dias meses y anyos
function ShF_SumarAFecha(Fecha: TDateTime; Anyos, Meses, Dias: Integer)
  : TDateTime;
// Mostrar ventanas de windows
function ShF_VentanaWindows(const Parametros: String): THandle;
// Abre XLS existente en la 1ra hoja
procedure ShF_OpenLeerXLS(Ruta: String; Excel: TExcelApplication;
  Libro: TExcelWorkbook; Hoja: TExcelWorksheet; VerHoja: Boolean;
  NameReport: String = '');
// Cierra XLS existente
procedure ShF_CloseLeerXLS(Ruta: String; Excel: TExcelApplication;
  Libro: TExcelWorkbook; Hoja: TExcelWorksheet; CerrarExcel: Boolean;
  NameReport: String = '');
// Crea XLS posicionandolo en la 1ra hoja y creando un numero de hojas
procedure ShF_OpenEscribirXLS(Ruta: String; NameReport: String;
  Excel: TExcelApplication; Libro: TExcelWorkbook; Hoja: TExcelWorksheet;
  NumHojas: Integer; NameHoja: String; VerHoja: Boolean);
// Guarda XLS creado
procedure ShF_CloseEscribirXLS(Ruta: String; NameReport: String;
  Excel: TExcelApplication; Libro: TExcelWorkbook; Hoja: TExcelWorksheet;
  NumHojas: Integer; CerrarExcel: Boolean);
// Comprime una carpeta en formato zip, indicando la carpeta y la ruta donde se quedara el zip
procedure ShF_CrearZIP(CarpetaAEmpacar, rutaZip: String);
// Copia un archivo a un Ftp
Function ShF_EnviarFileFTP(Ip, User, Pass, archivoorigen,
  archivodestino: String): Boolean;
// De toda una ruta retorna el nombre del archivo con o sin extension
Function ShF_NombreArchivo(Ruta: String; ConExt: Boolean): String;
// Borra carpeta y archivos
// function ShF_ALaPapelera(const Path: string): Boolean;
function ShF_ALaPapelera(Fichero: string): Boolean;

// Asigna la maxima prioridad a la aplicacion
procedure ShF_MaxPriori;
// Centra las Formas de la Aplicacion
procedure ShF_CentrarForm(Screen: TScreen);
// Lista los archivos TXT de un directorio en un string list
function ShF_ListaArchivosTXT(directorioPadre: string): TStringList;
// lista archivos segun el nombre
function ShF_ListaArchivosName(directorioPadre, Name: string): TStringList;
// Regresa cadena con el formato
function ShF_GenerarNombreConFecha(pref: String): String;
// Envio de correo
Function ShF_EnviarEmail(Para, CC, Titulo: String; Cuerpo: TCaption;
  ConAdjunto: Boolean = False; FileAdjunto: String = ''): Boolean;
// es numero una cadena
function ShF_IsNum(Cadena: string): Boolean;
// lista archivos segun extencion
function ShF_ListaArchivosExt(directorioPadre, Ext: string): TStringList;
// mover, copiar, borrar un archivo
procedure ShF_File(Archivo, DirDestino, Accion: String);
// nombre de usuario de windows
function ShF_GetUserName: String;
// lista archivos segun extencion   y Nombre
function ShF_ListaArchivosNameExt(directorioPadre, Ext, Name: string)
  : TStringList;
function ShF_GetFileDate(TheFileName: string): string;
function ShF_FindFile(const filespec: TFileName;
  attributes: Integer = 32): TStringList;
function ShF_Ceros(Vari: string; lng: Integer): string;
// Dir Temporal
function ShF_GetTempDirectory: String;
// iniciar session
function ShF_Session_ON(Session: TSession; pref: string;
  PrivateDir_NetFileDir: string = ''): String;
// Carrar Session
procedure ShF_Session_OFF(Session: TSession; SessionName: string;
  PrivateDir_NetFileDir: string = '');

procedure ShF_WordSetBold(RichE: TJvRichEdit; word: string);
procedure ShF_WordSetColor(RichE: TJvRichEdit; word: string; col: TColor);
function ShF_KillTask(ExeFileName: string): Integer;
function ShF_ProcessExists(ExeFileName: string): Boolean;
function ShF_WebIng(URL: PChar): Boolean;
function ShF_ConexionFTP(Host, User, Pass: String): Boolean;

function ShF_FicheroEnUso(Fichero: string): Boolean;

{ ************************************************
  ***************** SISTEMA CAPI ******************
  ************************************************* }
// Lista tablas de ini
function ShF_CAPI_ListaINI(Rubro: string): TStringList;
procedure ShF_CAPI_EstructuraUSB(Raiz: string);
function ShF_DateTimeToStr(Fecha: TDateTime): String;
function ShF_StrToDateTime(Fecha: String): TDateTime;
function ShF_CAPI_InsertarTMov(Fecha, FechaF, Tipo, Origen, Destino,
  FileName: String): Integer;
procedure ShF_CAPI_InsertarTListMov(n_mov: Integer;
  Control, Viv_Sel, Resul_V, Accion: String);
function ShF_CAPI_ResulxVivienda(jTRes_entrevista, Control, Viv_Sel: string)
  : String;

function ShF_CAPI_GeneraUVisita(TRes_entrevista, TablaDestino,
  VivHogPer: string): Boolean;
procedure ShF_CAPI_VerifFTP(Host, User, Pass, Dir: String);
function ShF_CAPI_DescripResulV(Resul_V: String): String;
function ShF_CAPI_AccionJE(Accion: String): String;
function ShF_CAPI_FechaUVisita(Usuario, Control, Viv_Sel: string): String;
function ShF_CAPI_Num_Vis_UVisita(Usuario, Control, Viv_Sel: string): String;
function ShF_CAPI_getUACT(Usuario: string): string;
procedure ShF_CAPI_setUACT(Usuario, Date: string);
Function ShF_CAPI_ValidarVersionTablasUSU(DirUsu: string): Boolean;
function ShF_CAPI_getVTab(Usuario: string): string;
procedure ShF_CAPI_setVTab(Usuario, Valor: string);

implementation

{ ************************************************
  ************** CONTROL SQL **********************
  ************************************************* }

procedure ShF_Pack(RutTabla: String);
var
  Database1: TDatabase;
  Tabla: TTable;
begin
  Database1 := TDatabase.Create(nil);
  Tabla := TTable.Create(nil);
  Tabla.TableName := RutTabla;
  if glbBanSession then
    Tabla.SessionName := glbSessionName;
  Tabla.Close;
  Tabla.Exclusive := True;
  Tabla.Open;
  Check(DbiPackTable(Database1.Handle, Tabla.Handle, '', '', True));
  Check(DbiPackTable(Tabla.DBHandle, Tabla.Handle, nil, szDBASE, True));
  Tabla.Close;
  Tabla.Exclusive := False;
  Tabla.Free;
  Database1.Free;
end;

Function ShF_BDE_OpenSQL(SQL: String; Qry: TQuery): Boolean;
var
  Res: Boolean;
begin
  Res := False;
  try
    try
      Qry.SQL.Clear;
      if glbBanSession then
        Qry.SessionName := glbSessionName;
      Qry.SQL.Add(SQL);
      Qry.Open;
      Qry.Active := True;
      if Qry.RecordCount > 0 then
      begin
        Res := True;
      end;
    except
      on E: Exception do
        ShF_ControlDeError
          ('[ShF.ShF_BDE_OpenSQL][' + E.Message + '] [' + SQL + ']');
    end;
  finally
    ShF_LiberarMemoria;
    Result := Res;
  end;
End;

procedure ShF_BDE_CloseSQL(Qry: TQuery);
begin
  try
    Qry.Active := False;
    // FreeAndNil(Qry);
    Qry.Close;
    Qry.Free;
    ShF_LiberarMemoria;
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_BDE_CloseSQL][' + E.Message + ']');
  end;
End;

procedure ShF_BDE_EjecutarSQL(SQL: String);
var
  Qry: TQuery;
begin
  Qry := TQuery.Create(nil);
  try
    try
      Qry.Close;
      Qry.SQL.Clear;
      if glbBanSession then
        Qry.SessionName := glbSessionName;
      Qry.SQL.Append(SQL);
      Qry.ExecSQL;
    except
      on E: Exception do
        ShF_ControlDeError
          ('[ShF.ShF_BDE_EjecutarSQL][' + E.Message + '] [' + SQL + ']');
    end;
  finally
    Qry.Close;
    Qry.Free;
    ShF_LiberarMemoria;
  end;
End;

Function ShF_BDE_OpenSelect(TQry: TQuery; Select: String; From: string;
  Where: String = ''; Condicion: String = ''): Boolean;
var
  SQL: String;
  Ban: Boolean;
begin
  Ban := False;
  try
    try
      SQL := 'SELECT ' + Select;
      SQL := SQL + ' FROM "' + From + '"';
      if Length(Trim(Where)) > 0 then
        SQL := SQL + ' WHERE ' + Where;
      if Length(Trim(Condicion)) > 0 then
        SQL := SQL + ' ' + Condicion;
      ShF_BDE_OpenSQL(SQL, TQry);
      if not TQry.Eof then
      begin
        Ban := True;
      end;
    except
      on E: Exception do
        ShF_ControlDeError
          ('[ShF.ShF_BDE_OpenSelect][' + E.Message + '] [' + SQL + ']');
    end;
  finally
    ShF_LiberarMemoria;
    Result := Ban;
  end;
end;

procedure ShF_ADO_OpenCONN(Dir: String; ADO: TADOConnection; Tipo: Integer);
var
  CONN: String;
begin
  if Tipo = 1 then
  begin
    CONN :=
      'Provider=MSDASQL.1;Persist Security Info=False;Extended Properties="DSN=FoxPro;UID=;';
    CONN := CONN + 'SourceDB=' + Dir + ';SourceType=DBF;Exclusive=No;';
    CONN := CONN + 'BackgroundFetch=Sí;Collate=Machine;"';
  end
  else if Tipo = 2 then
  begin
    CONN := 'Provider=VFPOLEDB.1;Data Source=' + Dir +
      ';Password="";Collating Sequence=MACHINE';
  end;
  ADO.ConnectionString := CONN;
  ADO.Connected := True;
End;

procedure ShF_ADO_OpenCONN_SQLSERV(Serv: String; DB: String;
  ADO: TADOConnection; ConAuten: Boolean = False;
  User: String = ''; Pass: String = '');
var
  CONN: String;
begin
  if not ConAuten then
  begin
    CONN :=
      'Provider=SQLNCLI.1;Integrated Security=SSPI;Persist Security Info=False;';
    CONN := CONN + 'Initial Catalog=' + DB + ';Data Source=' + Serv;
  end
  else
  begin
    // CONN:='Data Source='+Serv+';Initial Catalog='+DB+';User Id='+User+';Password='+Pass+';';
    CONN := 'Provider=SQLOLEDB.1;Password=' + Pass +
      ';Persist Security Info=True;User ID=' + User + ';';
    CONN := CONN + 'Initial Catalog=' + DB + ';Data Source=' + Serv;
  end;
  ADO.ConnectionString := CONN;
  ADO.Connected := True;
End;

procedure ShF_ADO_CloseCONN(ADO: TADOConnection);
begin
  ADO.Connected := False;
  ADO.Free;
End;

function ShF_ADO_OpenSQL(SQL: String; Qry: TADOQuery): Boolean;
var
  Res: Boolean;
begin
  Res := False;
  try
    try
      Res := False;
      Qry.SQL.Clear;
      Qry.SQL.Add(SQL);
      Qry.Open;
      Qry.Active := True;
      if Qry.RecordCount > 0 then
      begin
        Res := True;
      end;
    except
      on E: Exception do
        ShF_ControlDeError
          ('[ShF.ShF_ADO_OpenSQL][' + E.Message + '] [' + SQL + ']');
    end;
  Finally
    Result := Res;
  end;
End;

procedure ShF_ADO_CloseSQL(Qry: TADOQuery);
begin
  Qry.Active := False;
  Qry.Close;
  Qry.Free;
End;

procedure ShF_ADO_EjecutarSQL(SQL: String; Serv: String; DB: String;
  ConAuten: Boolean = False; User: String = ''; Pass: String = '');
var
  Qry: TADOQuery;
  CONN: TADOConnection;
begin
  Qry := TADOQuery.Create(nil);
  CONN := TADOConnection.Create(nil);
  try
    try
      if ConAuten then
      begin
        ShF_ADO_OpenCONN_SQLSERV(Serv, DB, CONN, True, User, Pass);
      end
      else
      begin
        ShF_ADO_OpenCONN_SQLSERV(Serv, DB, CONN);
      end;
      Qry.Connection := CONN;
      Qry.Close;
      Qry.SQL.Clear;
      Qry.SQL.Append(SQL);
      Qry.ExecSQL;
    except
      on E: Exception do
        ShF_ControlDeError
          ('[ShF.ShF_ADO_EjecutarSQL][' + E.Message + '] [' + SQL + ']');
    end;
  finally
    Qry.Close;
    Qry.Free;
    ShF_ADO_CloseCONN(CONN);
  end;
End;

Function ShF_ADO_OpenSelect(TQry: TADOQuery; Select: String; From: string;
  Where: String = ''; Condicion: String = ''): Boolean;
var
  SQL: String;
  Ban: Boolean;
begin
  Ban := False;
  try
    try
      SQL := 'SELECT ' + Select;
      SQL := SQL + ' FROM "' + From + '"';
      if Length(Trim(Where)) > 0 then
        SQL := SQL + ' WHERE ' + Where;
      if Length(Trim(Condicion)) > 0 then
        SQL := SQL + ' ' + Condicion;
      ShF_ADO_OpenSQL(SQL, TQry);
      if not TQry.Eof then
      begin
        Ban := True;
      end;
    except
      on E: Exception do
        ShF_ControlDeError
          ('[ShF.ShF_ADO_OpenSQL][' + E.Message + '] [' + SQL + ']');
    end;
  finally
    Result := Ban;
  end;
end;

function ShF_RowRSVacio(Qry: TQuery): Boolean;
var
  Band: Boolean;
  Cont: Integer;
begin
  Band := False;
  for Cont := 0 to Qry.FieldCount - 1 do
  begin
    if Trim(Qry.Fields[Cont].AsString) = '' then
    begin
      Band := True;
    end;
  end;
  Result := Band;
end;

function ShF_CamposNoVacios(Qry: TQuery): Integer;
var
  Cont, NumLlenos: Integer;
begin
  NumLlenos := 0;
  for Cont := 0 to Qry.FieldCount - 1 do
  begin
    if Trim(Qry.Fields[Cont].AsString) <> '' then
    begin
      NumLlenos := NumLlenos + 1;
    end;
  end;
  Result := NumLlenos;
end;

procedure ShF_MigrarTablaAllCamposString(oldTab: String; newTab: String;
  key: String);
var
  Cont: Integer;
  SQL, SQL_Insert, Campo: String;
  QryTab: TQuery;
begin
  QryTab := TQuery.Create(nil);
  SQL := 'SELECT *';
  SQL := SQL + ' FROM "' + oldTab + '"';
  SQL := SQL + ' WHERE Folio= ''' + key + '''';
  ShF_BDE_OpenSQL(SQL, QryTab);
  while not QryTab.Eof do
  begin
    SQL_Insert := 'INSERT INTO "' + newTab + '"';
    SQL_Insert := SQL_Insert + ' VALUES (';
    for Cont := 0 to QryTab.FieldCount - 1 do
    begin
      // Campo := ShF_CCampoRaros(AnsiReplaceStr(Trim(QryTab.Fields[Cont].AsString),'''', ' '),False);
      SQL_Insert := SQL_Insert + ' ''' + Campo + ''',';
    end;
    SQL_Insert := AnsiLeftStr(SQL_Insert, Length(SQL_Insert) - 1);
    SQL_Insert := SQL_Insert + ')';
    ShF_BDE_EjecutarSQL(SQL_Insert);
    QryTab.Next;
  end;
  ShF_BDE_CloseSQL(QryTab);
end;

procedure ShF_BDE_Insert_First(Tabla: string; Qry: TQuery);
var
  SQL, Campo: String;
  ContRows: Integer;
  DBTable001: TTable;
begin
  try
    DBTable001 := TTable.Create(nil);
    DBTable001.TableName := Tabla;
    if glbBanSession then
      DBTable001.SessionName := glbSessionName;
    DBTable001.Open;
    DBTable001.Active := True;
    DBTable001.Append;
    for ContRows := 0 to Qry.FieldCount - 1 do
    begin
      DBTable001.FieldByName(Qry.Fields[ContRows].FieldName).Value :=
        AnsiReplaceStr(Trim(Qry.Fields[ContRows].AsString), '''', ' ');
      Application.ProcessMessages;
    end;
    DBTable001.Post;
    DBTable001.Active := False;
    DBTable001.Close;
    DBTable001.Free;
    ShF_LiberarMemoria;
  except
    on E: Exception do
      ShF_ControlDeError
        ('[ShF.ShF_Insert_first][' + E.Message + '] [' + SQL + ']');
  end;
end;

procedure ShF_BDE_Insert_All(Tabla: string; Qry: TQuery);
var
  SQL, Campo: String;
  ContRows: Integer;
  DBTable001: TTable;
begin

  try
    DBTable001 := TTable.Create(nil);
    DBTable001.TableName := Tabla;
    if glbBanSession then
      DBTable001.SessionName := glbSessionName;
    DBTable001.FieldDefs.Assign(Qry.FieldDefs);
    glbTBachMove01 := TBatchMove.Create(nil);
    glbTBachMove01.Destination := DBTable001;
    glbTBachMove01.Source := Qry;
    glbTBachMove01.Execute;
    glbTBachMove01.Free;
    DBTable001.Free;
    ShF_LiberarMemoria;
  except
    on E: Exception do
      ShF_ControlDeError
        ('[ShF.ShF_Insert_all][' + E.Message + '] [' + SQL + ']');
  end;
end;

procedure ShF_Insert_ADO(Tabla: string; Qry: TADOQuery; Serv: String;
  DB: String);
var
  SQL, Campo: String;
  ContRows: Integer;
begin
  try
    SQL := 'INSERT INTO ' + Tabla + ' VALUES (';
    for ContRows := 0 to Qry.FieldCount - 1 do
    begin
      Campo := AnsiReplaceStr(Trim(Qry.Fields[ContRows].AsString), '''', ' ');
      if Qry.Fields[ContRows].DataType = ftDate then
      begin
        SQL := SQL + ' CAST(''' + Campo + ''' AS DATE),';
      end
      else if (Qry.Fields[ContRows].DataType = ftInteger) or
        (Qry.Fields[ContRows].DataType = ftFloat) then
      begin
        if Trim(Campo) = '' then
        begin
          SQL := SQL + ' NULL,';
        end
        else
        begin
          SQL := SQL + ' ' + Campo + ',';
        end;
      end
      else
      begin
        SQL := SQL + ' ''' + Campo + ''',';
      end;
      Application.ProcessMessages;
    end;
    SQL := AnsiLeftStr(SQL, Length(SQL) - 1) + ')';
    ShF_ADO_EjecutarSQL(SQL, Serv, DB);
  except
    on E: Exception do
      ShF_ControlDeError
        ('[ShF.ShF_Insert_ADO][' + E.Message + '] [' + SQL + ']');
  end;
end;

function ShF_Where(Qry: TQuery): String;
var
  SQL, Campo: String;
  ContRows: Integer;
begin
  SQL := '';
  try
    try
      for ContRows := 0 to Qry.FieldCount - 1 do
      begin
        Campo := AnsiReplaceStr(Trim(Qry.Fields[ContRows].AsString), '''', ' ');
        SQL := SQL + ' ' + Qry.Fields[ContRows].FieldName + ' =';
        if Qry.Fields[ContRows].DataType = ftDate then
        begin
          SQL := SQL + ' Cast(''' + Campo + ''' as Date) AND';
        end
        else if (Qry.Fields[ContRows].DataType = ftInteger) or
          (Qry.Fields[ContRows].DataType = ftFloat) then
        begin
          SQL := SQL + ' ' + Campo + ' AND';
        end
        else
        begin
          SQL := SQL + ' ''' + Campo + ''' AND';
        end;
        Application.ProcessMessages;
      end;
      SQL := Trim(AnsiLeftStr(SQL, Length(SQL) - 3));
    except
      on E: Exception do
        ShF_ControlDeError('[ShF.ShF_Where][' + E.Message + '] [' + SQL + ']');
    end;
  finally
    Result := SQL;
  end;
end;

procedure ShF_ActualziarRegistros(Origen, Destino, LLave: String);
var
  QryA: TQuery;
  SQL: String;
begin
  QryA := TQuery.Create(nil);
  try
    try
      SQL := 'DELETE FROM "' + Destino + '"';
      SQL := SQL + ' WHERE ' + LLave;
      ShF_BDE_EjecutarSQL(SQL);
      if ShF_BDE_OpenSelect(QryA, '*', Origen, LLave) then
      begin
        ShF_BDE_Insert_All(Destino, QryA);
      end;
      ShF_Pack(Destino);
    except
      on E: Exception do
        ShF_ControlDeError
          ('[ShF.ShF_ActualziarRegistros][' + E.Message + '] [' + SQL +
            '][' + Origen + '][' + Destino + '][' + LLave + ']');
    end;
  finally
    ShF_BDE_CloseSQL(QryA);
    ShF_LiberarMemoria;
  end;
end;

procedure ShF_CrearTabla(NameTabla: string; Qry: TQuery; Tipo: TTableType);
var
  Tabla: TTable;
  Cont: Integer;
begin
  Tabla := TTable.Create(Application);
  try
    try
      Tabla.TableName := NameTabla;
      if glbBanSession then
        Tabla.SessionName := glbSessionName;
      Tabla.FieldDefs.Assign(Qry.FieldDefs);
      Tabla.TableType := Tipo;
      Tabla.CreateTable;
    except
      on E: Exception do
        ShF_ControlDeError('[ShF.ShF_CrearTabla][' + E.Message + '] [' +
            NameTabla + ']');
    end;
  finally
    Tabla.Free;
    ShF_LiberarMemoria;
  end;
end;

procedure ShF_TQryToDBF(NameTabla: string; Qry: TQuery; Tipo: TTableType);
var
  Tabla: TTable;
  Cont: Integer;
begin
  Tabla := TTable.Create(Application);
  try
    try
      With Tabla do
      begin
        Active := False;
        TableName := NameTabla;
        TableType := Tipo;
        FieldDefs.Clear;
      end;
      if glbBanSession then
        Tabla.SessionName := glbSessionName;
      for Cont := 0 to Qry.FieldCount - 1 do
      begin
        if Qry.Fields[Cont].DataType = ftString then
        begin
          Tabla.FieldDefs.Add(Qry.Fields[Cont].FieldName,
            Qry.Fields[Cont].DataType, Qry.Fields[Cont].DataSize, False);
        end
        else
        begin
          Tabla.FieldDefs.Add(Qry.Fields[Cont].FieldName,
            Qry.Fields[Cont].DataType, 0, False);
        end;
        Tabla.CreateTable;
        Application.ProcessMessages;
      end;
      Qry.First;
      ShF_BDE_Insert_All(NameTabla, Qry);
    except
      on E: Exception do
        ShF_ControlDeError('[ShF.ShF_TQryToDBF][' + E.Message + '] [' +
            NameTabla + ']');
    end;
  finally
    Tabla.Destroy;
    ShF_LiberarMemoria;
  end;
end;

function ShF_CuantosRegistros(Tablas: String): String;
var
  Cuantos: String;
  Qry: TQuery;
begin
  Cuantos := '';
  Qry := TQuery.Create(nil);
  if ShF_BDE_OpenSelect(Qry, 'Count(*) as Cuantos', Tablas) then
  begin
    Cuantos := Trim(UpperCase(Qry.FieldByName('Cuantos').AsString));
  end;
  ShF_BDE_CloseSQL(Qry);
  Result := Cuantos;
end;

{ ************************************************
  **************** UTILERIAS **********************
  ************************************************* }

procedure ShF_LiberarMemoria;
begin
  if Win32Platform = VER_PLATFORM_WIN32_NT then
    SetProcessWorkingSetSize(GetCurrentProcess, $FFFFFFFF, $FFFFFFFF);
end;

procedure ShF_FileError(Nombre, Error: String);
var
  F: TextFile;
begin
  AssignFile(F, Nombre);
  if FileExists(Nombre) then
    Append(F)
  else
    Rewrite(F);
  WriteLn(F, Error);
  CloseFile(F);
end;

procedure ShF_ControlDeError(Error: String);
begin
  Inc(glbContError);
  if glbMostrarErrores then
    ShowMessage('[' + IntToStr(glbContError) + ']' + Error);
  ShF_FileError(glbNameFileError, '[' + IntToStr(glbContError) + ']' + Error);
end;

function ShF_GetPosiStrig(s: string; Indice: Byte; Separador: string): string;
var
  i: Integer;
  tmp: string;
begin
  i := 1;
  while i <= Indice do
  begin
    Delete(s, 1, Pos(Separador, s));
    Inc(i);
  end;
  if Pos(Separador, s) <> 0 then
    tmp := Copy(s, 1, Pos(Separador, s) - 1)
  else
    tmp := s;
  if Length(tmp) = 0 then
    tmp := '';
  Result := tmp;
end;

Function ShF_CuantosCharHay(Cadena: String; Caracter: Char): Integer;
var
  Pos, Veces: Integer;
  Nombre: String;
begin
  Veces := 0;
  for Pos := 0 to Length(Cadena) - 1 do
    if Cadena[Pos] = Caracter then
      Inc(Veces);
  Result := Veces;
end;

Function ShF_DirUp(Ruta: String): String;
var
  Pos, Veces: Integer;
  Nombre: String;
begin
  Veces := ShF_CuantosCharHay(Ruta, '\');
  Nombre := AnsiLeftStr(Ruta, Length(Ruta) - Length(ShF_GetPosiStrig(Ruta,
        Veces, '\')));
  Result := AnsiLeftStr(Nombre, Length(Nombre) - 1);
end;

function ShF_BorrarCarpeta(const rutaCarpeta: string): Boolean;
var
  FileInfo: TShFileOpStruct;
begin
  FileInfo.Wnd := 0;
  FileInfo.wFunc := FO_DELETE;
  FileInfo.pFrom := PChar(rutaCarpeta);
  FileInfo.pTo := nil;
  FileInfo.fFlags := FOF_NOERRORUI or FOF_NOCONFIRMATION;
  ShFileOperation(FileInfo);
  Result := True;
end;

function ShF_LimpiarString(Campo: String; Acentos: Boolean): String;
var
  Cont: Integer;
  Caracter: Char;
begin
  Cont := 1;
  while (Cont <= Length(Campo)) do
  begin
    Caracter := Campo[Cont];
    if StrScan(
      '0123456789 áÁéÉíÍóÓúÚabcdfeghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ:´.,(){}[]/\_-¿?¡!%#$&=|+*''', Caracter) = nil then
    begin
      Campo := AnsiLeftStr(Campo, Cont - 1) + AnsiRightStr(Campo,
        Length(Campo) - Cont);
    end
    else
    begin
      if (StrScan('áÁéÉíÍóÓúÚ', Caracter) <> nil) and (Acentos) then
      begin
        case Caracter of
          'á':
            Caracter := 'a';
          'Á':
            Caracter := 'A';
          'é':
            Caracter := 'e';
          'É':
            Caracter := 'E';
          'í':
            Caracter := 'i';
          'Í':
            Caracter := 'I';
          'ó':
            Caracter := 'o';
          'Ó':
            Caracter := 'O';
          'ú':
            Caracter := 'u';
          'Ú':
            Caracter := 'U';
        end;
        Campo := AnsiLeftStr(Campo, Cont - 1) + Caracter + AnsiRightStr(Campo,
          Length(Campo) - Cont);
        Inc(Cont);
      end
      else
      begin
        Inc(Cont);
      end;
    end;
  end;
  Result := Campo;
end;

{ function ShF_ListaDirectorios(directorioPadre: string ; ComenzarCon: string = ''):TStringList;
  var
  sr: TSearchRec;
  begin
  Result := TStringList.Create;
  Result.Clear;
  if FindFirst(directorioPadre + '*', faDirectory, sr) = 0 then
  repeat
  if (sr.Attr = faDirectory) then
  if ( sr.Name <> '.' ) and ( sr.Name <> '..' ) then begin
  if Length(Trim(ComenzarCon))= 0 then begin
  Result.Add(sr.Name);
  end else begin
  if UpperCase(Trim(ComenzarCon))= UpperCase( AnsiLeftStr( Trim(sr.Name), length( Trim(ComenzarCon) ) ) ) then
  Result.Add(sr.Name);
  end;
  end;
  until FindNext(sr) <> 0;
  FindClose(sr);
  end;
}

function ShF_ListaDirectorios(directorioPadre: string;
  ComenzarCon: string = ''): TStringList;
var
  Directorio: TSearchRec;
  iResultado: Integer;
begin
  if directorioPadre[Length(directorioPadre)] <> '\' then
    directorioPadre := directorioPadre + '\';

  Result := TStringList.Create;
  Result.Clear;

  iResultado := FindFirst(directorioPadre + '*.*', FaAnyfile, Directorio);
  while iResultado = 0 do
  begin
    // ¿Es un directorio y hay que entrar en él?
    if (Directorio.Attr and faDirectory = faDirectory) then
    begin
      if (Directorio.Name <> '.') and (Directorio.Name <> '..') then
      begin
        if Length(Trim(ComenzarCon)) = 0 then
        begin
          Result.Add(Directorio.Name);
        end
        else
        begin
          if UpperCase(Trim(ComenzarCon)) = UpperCase
            (AnsiLeftStr(Trim(Directorio.Name), Length(Trim(ComenzarCon))))
            then
            Result.Add(Directorio.Name);
        end;
      end;
    end;
    iResultado := FindNext(Directorio);
  end;
  FindClose(Directorio);

end;

function ShF_ListaArchivos(directorioPadre: string): TStringList;
var
  sr: TSearchRec;
begin
  Result := TStringList.Create;
  if FindFirst(directorioPadre + '\*', FaAnyfile, sr) = 0 then
    repeat
      if (sr.Attr and faDirectory = 0) or (sr.Name <> '.') and
        (sr.Name <> '..') then
      begin
        Result.Add(sr.Name);
      end;
    until FindNext(sr) <> 0;
    FindClose(sr);
end;

Procedure ShF_InicializarStringGrid(SGrid: TStringGrid);
var
  Cont: Integer;
begin
  for Cont := 0 to SGrid.RowCount - 1 do
    SGrid.Rows[Cont].Clear;
  SGrid.RowCount := 1;
  for Cont := 0 to SGrid.ColCount - 1 do
    SGrid.Cols[Cont].Clear;
  SGrid.RowCount := 1;
  SGrid.ColCount := 1;
end;

{ function ShF_UnZipFile(DirArchivo: string; DirDestino: String) : boolean;
  var
  FileUnZip: TUnZip;
  begin
  try try
  FileUnZip := TUnZip.Create(nil);
  FileUnZip.OverwriteMode := omOverwrite;
  FileUnZip.ArchiveFile := DirArchivo;
  FileUnZip.ConfirmOverwrites := True;
  FileUnZip.RecurseDirs := True;
  FileUnZip.FileSpec.Clear();
  FileUnZip.FileSpec.Add('*.*');
  FileUnZip.ExtractDir := DirDestino;
  FileUnZip.Extract();
  except
  on E: Exception do
  ShF_ControlDeError('[ShF.ShF_UnZipFile]['+ E.Message +'] ['+ DirArchivo+']['+ DirDestino +']');
  end;
  Finally
  Result:= True;
  end;
  end; }

function ShF_UnZipFile(DirArchivo: string; DirDestino: String): Boolean;
var
  FileUnZip: TAbUnZipper;
  desempaco : Boolean;
begin
  try
    try
      desempaco := True;
      FileUnZip := TAbUnZipper.Create(nil);
      FileUnZip.FileName := DirArchivo;
      FileUnZip.BaseDirectory := DirDestino;
      FileUnZip.ExtractOptions := [eoCreateDirs, eoRestorePath];
      FileUnZip.ExtractFiles('*.*');
      FileUnZip.Free;

    except
      on E: Exception do begin
        desempaco := False;
        ShF_ControlDeError('[ShF.ShF_UnZipFile][' + E.Message + '] [' +
            DirArchivo + '][' + DirDestino + ']');
      end;
    end;
  Finally
    Result := desempaco;
  end;
end;

procedure ShF_CrearZIP(CarpetaAEmpacar, rutaZip: String);
var
  WinZip: TAbZipKit;
begin
  try
    Screen.Cursor := crHourGlass;
    WinZip := TAbZipKit.Create(nil);
    WinZip.StoreOptions := [soRecurse];
    WinZip.FileName := rutaZip;
    WinZip.BaseDirectory := CarpetaAEmpacar;
    WinZip.AddFiles('*.*', 0);
    WinZip.CloseArchive;
    WinZip.Save;
    WinZip.Free;
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_CrearZIP][' + E.Message + '] [' +
          CarpetaAEmpacar + '][' + rutaZip + ']');
  end;
  Screen.Cursor := crDefault;
end;

function ShF_CopiaTodo(Origen, Destino: String): LongInt;
var
  F: TShFileOpStruct;
  sOrigen, sDestino: String;
begin
  sOrigen := Origen + #0;
  sDestino := Destino + #0;
  with F do
  begin
    Wnd := Application.Handle;
    wFunc := FO_COPY;
    pFrom := @sOrigen[1];
    pTo := @sDestino[1];
    fFlags := FOF_ALLOWUNDO or FOF_NOCONFIRMATION
  end;
  Result := ShFileOperation(F);
end;

function ShF_LeerINI(Dir, Secc, Vari: string): String;
var
  R: String;
begin
  with TIniFile.Create(Dir) do
    try
      R := ReadString(Secc, Vari, '#NoCampo#');
    finally
      Free;
    end;
  ShF_LeerINI := R;
end;

function ShF_EscribirINI(Dir, Secc, Vari, Valor: string): Boolean;
begin
  with TIniFile.Create(Dir) do
    try
      WriteString(Secc, Vari, Valor);
    finally
      Free;
    end;
  ShF_EscribirINI := True;
end;

Function ShF_CantidadEnLetra(curCantidad: Currency; MonNal: Integer): String;
var
  i: Integer;
  Cantidad, Centavos: Currency;
  BloqueCero, NumeroBloques, Digito: Byte;
  PrimerDigito, SegundoDigito, TercerDigito: Byte;
  Resultado, Temp, strCentavos, Bloque: String;
  Unidades: Array [0 .. 28] of String;
  Decenas: Array [0 .. 8] of String;
  Centenas: Array [0 .. 8] of String;
begin
  Unidades[0] := 'UN';
  Unidades[1] := 'DOS';
  Unidades[2] := 'TRES';
  Unidades[3] := 'CUATRO';
  Unidades[4] := 'CINCO';
  Unidades[5] := 'SEIS';
  Unidades[6] := 'SIETE';
  Unidades[7] := 'OCHO';
  Unidades[8] := 'NUEVE';
  Unidades[9] := 'DIEZ';
  Unidades[10] := 'ONCE';
  Unidades[11] := 'DOCE';
  Unidades[12] := 'TRECE';
  Unidades[13] := 'CATORCE';
  Unidades[14] := 'QUINCE';
  Unidades[15] := 'DIESISEIS';
  Unidades[16] := 'DIESISIETE';
  Unidades[17] := 'DIESIOCHO';
  Unidades[18] := 'DIESINUEVE';
  Unidades[19] := 'VEINTE';
  Unidades[20] := 'VEINTIUNO';
  Unidades[21] := 'VEINTIDOS';
  Unidades[22] := 'VEINTITRES';
  Unidades[23] := 'VEINTICUATRO';
  Unidades[24] := 'VEINTICINCO';
  Unidades[25] := 'VEINTISEIS';
  Unidades[26] := 'VEINTISIETE';
  Unidades[27] := 'VEINTIOCHO';
  Unidades[28] := 'VEINTINUEVE';
  Decenas[0] := 'DIEZ';
  Decenas[1] := 'VEINTE';
  Decenas[2] := 'TREINTA';
  Decenas[3] := 'CUARENTA';
  Decenas[4] := 'CINCUENTA';
  Decenas[5] := 'SESENTA';
  Decenas[6] := 'SETENTA';
  Decenas[7] := 'OCHENTA';
  Decenas[8] := 'NOVENTA';
  Centenas[0] := 'CIENTO';
  Centenas[1] := 'DOSCIENTOS';
  Centenas[2] := 'TRESCIENTOS';
  Centenas[3] := 'CUATROCIENTOS';
  Centenas[4] := 'QUINIENTOS';
  Centenas[5] := 'SEISCIENTOS';
  Centenas[6] := 'SETECIENTOS';
  Centenas[7] := 'OCHOCIENTOS';
  Centenas[8] := 'NOVECIENTOS';

  Cantidad := Trunc(curCantidad);
  Centavos := (curCantidad - Cantidad) * 100;
  NumeroBloques := 1;
  Repeat
    PrimerDigito := 0;
    SegundoDigito := 0;
    TercerDigito := 0;
    Bloque := '';
    BloqueCero := 0;
    For i := 1 To 3 do
    begin
      Digito := Round(Cantidad) Mod 10;
      If Digito <> 0 Then
      begin
        Case i of
          1:
            begin
              Bloque := ' ' + Unidades[Digito - 1];
              PrimerDigito := Digito;
            end;
          2:
            begin
              If Digito <= 2 Then
              begin
                Bloque := ' ' + Unidades[(Digito * 10 + PrimerDigito - 1)];
              end
              Else
              begin
                If PrimerDigito <> 0 then
                  Temp := ' Y'
                else
                  Temp := '';
                Bloque := ' ' + Decenas[Digito - 1] + Temp + Bloque;
              End;
              SegundoDigito := Digito;
            end;
          3:
            begin
              If (Digito = 1) and (PrimerDigito = 0) and (SegundoDigito = 0)
                then
                Temp := 'CIEN'
              else
                Temp := Centenas[Digito - 1];
              Bloque := ' ' + Temp + Bloque;
              TercerDigito := Digito;
            end;
        End;
      end
      Else
      begin
        BloqueCero := BloqueCero + 1;
      End;
      Cantidad := Int(Cantidad / 10);
      If Cantidad = 0 Then
      begin
        Break;
      End;
    end;
    Case NumeroBloques of
      1:
        Resultado := Bloque;
      2:
        begin
          if BloqueCero = 3 then
            Temp := ''
          else
            Temp := ' MIL';
          Resultado := Bloque + Temp + Resultado;
        end;
      3:
        begin
          If (PrimerDigito = 1) and (SegundoDigito = 0) and (TercerDigito = 0)
            then
            Temp := ' MILLON'
          else
            Temp := ' MILLONES';
          Resultado := Bloque + Temp + Resultado;
        end;
    End;
    NumeroBloques := NumeroBloques + 1;
  Until Cantidad = 0;
  case MonNal of
    0:
      begin
        If curCantidad > 1 then
          Temp := ' CENTAVOS ***'
        else
          Temp := ' CENTAVO ***';
        ShF_CantidadEnLetra := Resultado + Temp;
      end;
    1:
      begin
        If curCantidad > 1 then
          Temp := ' PESOS '
        else
          Temp := ' PESO ';
        if Centavos = 0 then
          strCentavos := ''
        else
          strCentavos := 'CON ' + ShF_CantidadEnLetra(Centavos, 0);
        ShF_CantidadEnLetra := 'SON: *** ' + Resultado + Temp + strCentavos;
      end;
    2:
      begin
        If curCantidad > 1 then
          Temp := ' DLLS '
        else
          Temp := ' DOLAR ';
        if Centavos = 0 then
          strCentavos := ''
        else
          strCentavos := 'CON ' + ShF_CantidadEnLetra(Centavos, 0);
        ShF_CantidadEnLetra := 'SON: *** ' + Resultado + Temp + strCentavos;
      end;
  end;
End;

function ShF_FindInStr(Text, Destination: String;
  CaseSensitive: Boolean): Boolean;
begin
  if CaseSensitive then
    Destination := LowerCase(Destination);
  Result := Pos(Text, Destination) > 0;
end;

function ShF_FindInTextFile(Text, Filepath: String;
  CaseSensitive: Boolean): Boolean;
var
  AFile: TStringList;
begin
  if FileExists(Filepath) then
  begin
    AFile := TStringList.Create;
    try
      AFile.LoadFromFile(Filepath);
      Result := ShF_FindInStr(Text, AFile.Text, CaseSensitive);
    finally
      AFile.Free;
    end;
  end
  else
  begin
    Result := False;
  end;
end;

function ShF_Strip(s: string; tildes: Boolean): string;
var
  i: Byte;
  t: string;
begin
  if s = '' then
    ShF_Strip := 'null'
  else
  begin
    t := '';
    if tildes = True then
      for i := 1 to Length(s) do
      begin
        case s[i] of
          'à', 'á':
            s[i] := 'a';
          'Á', 'À':
            s[i] := 'A';
          'è', 'é':
            s[i] := 'e';
          'É', 'È':
            s[i] := 'E';
          'ì', 'í':
            s[i] := 'i';
          'Í', 'Ì':
            s[i] := 'I';
          'ò', 'ó':
            s[i] := 'o';
          'Ó', 'Ò':
            s[i] := 'O';
          'ù', 'ú':
            s[i] := 'u';
          'Ú', 'Ù':
            s[i] := 'U';
          '\', '/', ':', '*', '?', '"', '|', '<', '>', '&':
            s[i] := ' ';
        end;
        t := t + s[i];
      end
      else
        for i := 1 to Length(s) do
        begin
          case s[i] of
            '\', '/', ':', '*', '?', '"', '|', '<', '>', '&':
              s[i] := ' ';
          end;
          t := t + s[i];
        end;
    ShF_Strip := t;
  end;
end;

function ShF_ListaArchivosZip(directorioPadre: string): TStringList;
{ Saca todos los nombres de los archivos zip en un directorio }
var
  sr: TSearchRec;
begin
  Result := TStringList.Create;
  if FindFirst(directorioPadre + '\*', FaAnyfile, sr) = 0 then
    repeat
      if (sr.Attr and faDirectory = 0) or (sr.Name <> '.') and
        (sr.Name <> '..') then
      begin
        if UpperCase(Trim(AnsiRightStr(sr.Name, Length(sr.Name) - PosEx('.',
                sr.Name)))) = 'ZIP' then
        begin { Si es *.zip }
          Result.Add(AnsiLeftStr(sr.Name, PosEx('.', sr.Name) - 1));
          { Agregarlo en el TStringList }
        end;
      end;
    until FindNext(sr) <> 0;
    FindClose(sr);
end;

function ShF_12HrPunto(): String;
var
  all, h, m, s, x: String;
begin
  all := FormatDateTime('hh:mm:ss am/pm', Now());
  h := AnsiLeftStr(all, 2);
  m := AnsiMidStr(all, 4, 2);
  s := AnsiMidStr(all, 7, 2);
  x := AnsiRightStr(all, 2);
  if x = 'am' then
    x := 'a.m.';
  if x = 'pm' then
    x := 'p.m.';
  ShF_12HrPunto := h + ':' + m + ':' + s + ' ' + x;
end;

function ShF_Fecha_12HrsSinPunto(Tipo: String): String;
var
  Date, dd, mm, aaaa, hh, min, ss, Resultado: String;
begin
  try
    try
      Resultado := '';
      Date := FormatDateTime('ddmmyyyy_HHmmss', Now());
      dd := AnsiLeftStr(Date, 2);
      mm := AnsiMidStr(Date, 3, 2);
      aaaa := AnsiMidStr(Date, 5, 4);
      hh := AnsiMidStr(Date, 10, 2);
      min := AnsiMidStr(Date, 12, 2);
      ss := AnsiMidStr(Date, 14, 2);
      if Tipo = 'F' then
      begin
        Resultado := dd + '/' + mm + '/' + aaaa;
      end
      else if Tipo = 'H' then
      begin
        Resultado := hh + ':' + min + ':' + ss;
      end
      else if Tipo = 'T' then
      begin
        Resultado := dd + '/' + mm + '/' + aaaa + ' ' + hh + ':' + min + ':' +
          ss;
      end;
    except
      on E: Exception do
        ShF_ControlDeError('[ShF.ShF_Fecha_12HrsSinPunto][' + E.Message + ']');
    end;
  finally
    Result := Resultado;
  end;
end;

function ShF_Edad(FechaNacimiento: string): Integer;
var
  iTemp, iTemp2, Nada: word;
  Fecha: TDate;
begin
  Fecha := StrToDate(FechaNacimiento);
  DecodeDate(Date, iTemp, Nada, Nada);
  DecodeDate(Fecha, iTemp2, Nada, Nada);
  if FormatDateTime('mmdd', Date) < FormatDateTime('mmdd', Fecha) then
    Result := iTemp - iTemp2 - 1
  else
    Result := iTemp - iTemp2;
end;

function ShF_DiasTranscurridos(FechaIni, FechaFin: string): Integer;
var
  dTemp: TDate;
begin
  dTemp := StrToDate(FechaFin);
  Result := Trunc(dTemp - StrToDate(FormatDateTime(FechaIni, dTemp))) + 1;
end;

function ShF_SumarAFecha(Fecha: TDateTime; Anyos, Meses, Dias: Integer)
  : TDateTime;
var
  d, m, a: Integer;
  d2, m2, a2: word;
begin
  DecodeDate(Fecha, a2, m2, d2);
  a := a2 + Anyos;
  m := m2 + Meses - 1;
  d := d2 + Dias - 1;
  if m > 0 then
  begin
    a := a + (m div 12);
    m := (m mod 12) + 1;
  end
  else
  begin
    m := -m;
    a := a - (m div 12) - 1;
    m := 13 - (m mod 12);
  end;
  Result := EncodeDate(a, m, 1) + d;
end;

function ShF_VentanaWindows(const Parametros: String): THandle;
begin
  Result := ShellExecute(Application.MainForm.Handle, nil,
    PChar('C:\WINDOWS\system32\rundll32.exe'), PChar(Parametros), nil,
    SW_SHOW);
end;

procedure ShF_OpenLeerXLS(Ruta: String; Excel: TExcelApplication;
  Libro: TExcelWorkbook; Hoja: TExcelWorksheet; VerHoja: Boolean;
  NameReport: String = '');
var
  xlsFile: String;
begin
  if Length(Trim(NameReport)) > 0 then
    xlsFile := Ruta + '\' + Trim(NameReport)
  else
    xlsFile := Ruta;
  Excel.Connect;
  Excel.Visible[LOCALE_USER_DEFAULT] := False;
  Excel.Workbooks.Open(xlsFile, False, False, emptyparam, '', False, False,
    emptyparam, emptyparam, False, False, emptyparam, False, emptyparam, False,
    0);
  Libro.ConnectTo(Excel.ActiveWorkBook);
  Hoja.ConnectTo(Excel.Sheets.Item[1] as _Worksheet);
end;

procedure ShF_CloseLeerXLS(Ruta: String; Excel: TExcelApplication;
  Libro: TExcelWorkbook; Hoja: TExcelWorksheet; CerrarExcel: Boolean;
  NameReport: String = '');
var
  xlsFile: String;
begin
  if Length(Trim(NameReport)) > 0 then
    xlsFile := Ruta + '\' + Trim(NameReport)
  else
    xlsFile := Ruta;
  Excel.DisplayAlerts[LOCALE_USER_DEFAULT] := False;
  Libro.Close(True, Trim(xlsFile), emptyparam);
  Hoja.Disconnect;
  Libro.Disconnect;
  Excel.Disconnect;
  if CerrarExcel then
    Excel.Quit;
  Hoja.Free;
  Libro.Free;
  Excel.Free;
end;

procedure ShF_OpenEscribirXLS(Ruta: String; NameReport: String;
  Excel: TExcelApplication; Libro: TExcelWorkbook; Hoja: TExcelWorksheet;
  NumHojas: Integer; NameHoja: String; VerHoja: Boolean);
var
  xlsFile: String;
  Cont: Integer;
begin
  xlsFile := Ruta + '\' + Trim(NameReport);
  Excel.Connect;
  if VerHoja then
    Excel.Visible[LOCALE_USER_DEFAULT] := True
  else
    Excel.Visible[LOCALE_USER_DEFAULT] := False;
  Excel.NewWorkbook.Add(xlsFile, emptyparam, 1, emptyparam);
  Excel.Workbooks.Add(emptyparam, LOCALE_USER_DEFAULT);
  Libro.ConnectTo(Excel.ActiveWorkBook);
  while Libro.Sheets.Count > 1 do
  begin
    Hoja.ConnectTo(Libro.Sheets.Item[Libro.Sheets.Count] as _Worksheet);
    Hoja.Delete;
    Hoja.Disconnect;
  end;
  for Cont := 2 to NumHojas do
    Libro.Sheets.Add(emptyparam, emptyparam, 1, emptyparam, 0);
  Hoja.ConnectTo(Libro.Sheets.Item[1] as _Worksheet);
end;

procedure ShF_CloseEscribirXLS(Ruta: String; NameReport: String;
  Excel: TExcelApplication; Libro: TExcelWorkbook; Hoja: TExcelWorksheet;
  NumHojas: Integer; CerrarExcel: Boolean);
var
  xlsFile: String;
begin
  xlsFile := Ruta + '\' + Trim(NameReport);
  { Hoja.Disconnect;
    while Libro.Sheets.Count > NumHojas  do begin
    Hoja.ConnectTo(Libro.Sheets.Item[Libro.Sheets.Count] as _Worksheet);
    Hoja.Delete;
    Hoja.Disconnect;
    end; }
  Excel.DisplayAlerts[LOCALE_USER_DEFAULT] := False;
  Libro.Close(True, xlsFile, emptyparam);
  Hoja.Disconnect;
  Libro.Disconnect;
  Excel.Disconnect;
  Excel.Quit;
end;
{
  procedure ShF_CrearZIP(CarpetaAEmpacar, rutaZip:String);
  var
  WinZip: TZip;
  begin
  try
  Screen.Cursor:=crHourGlass;
  WinZip := TZip.Create(nil);
  WinZip.ArchiveFile := rutaZip;
  WinZip.DateAttribute := daFileDate;
  WinZip.StoredDirNames := sdRelative;
  WinZip.CompressMethod := cmDeflate;
  WinZip.StoreEmptySubDirs := False;
  WinZip.ExcludeSpec.Clear();
  WinZip.FileSpec.Clear();
  WinZip.FileSpec.Add(CarpetaAEmpacar + '\*.*');
  WinZip.SetAttributeEx(fsZeroAttr, False);
  WinZip.SetAttributeEx(fsArchive, False);
  WinZip.SetAttributeEx(fsDirectory, False);
  WinZip.SetAttributeEx(fsHidden, False);
  WinZip.SetAttributeEx(fsReadOnly, False);
  WinZip.SetAttributeEx(fsSysFile, False);
  WinZip.Compress();
  WinZip.Free;
  except
  on E: Exception do
  ShF_ControlDeError('[ShF.ShF_CrearZIP]['+ E.Message +'] ['+ CarpetaAEmpacar+']['+ rutaZip +']');
  end;
  Screen.Cursor:=crDefault ;
  end;
  }

Function ShF_EnviarFileFTP(Ip, User, Pass, archivoorigen,
  archivodestino: String): Boolean;
var
  Archivo, BytesE, BytesR: Integer;
  FTP: TIdFTP;
  flag: Boolean;
begin
  flag := False;
  try
    try
      BytesR := 0;
      FTP := TIdFTP.Create(nil);
      FTP.Username := User;
      FTP.Password := Pass;
      FTP.Host := Ip;
      FTP.Connect;
      If FTP.Connected then
      begin
        FTP.BeginWork(wmWrite);
        Try
          FTP.Put(archivoorigen, archivodestino, False);
          BytesR := FTP.Size(archivodestino);
          FTP.EndWork(wmWrite);
          FTP.Disconnect;
        Finally
          Archivo := FileOpen(archivoorigen, 0);
          BytesE := getfilesize(Archivo, nil);
          FileClose(Archivo);
          If BytesR >= BytesE then
          begin
            flag := True;
          end;
        End;
      end;
    except
      on E: Exception do
        ShF_ControlDeError('[ShF.ShF_EnviarFileFTP][' + E.Message + '] [' +
            archivoorigen + '][' + archivodestino + ']');
    end;
  finally
    Result := flag;
  end;
end;

Function ShF_NombreArchivo(Ruta: String; ConExt: Boolean): String;
var
  Pos, Veces: Integer;
  Nombre: String;
begin
  Veces := ShF_CuantosCharHay(Ruta, '\');
  if ConExt then
    Nombre := ShF_GetPosiStrig(Ruta, Veces, '\')
  else
    Nombre := AnsiLeftStr(ShF_GetPosiStrig(Ruta, Veces, '\'),
      PosEx('.', ShF_GetPosiStrig(Ruta, Veces, '\')) - 1);
  // Nombre := AnsiLeftStr(Ruta, Length(Ruta)- Length(ShF_GetPosiStrig(Ruta ,veces,'\')) );
  Result := Nombre;
end;

{
  function ShF_ALaPapelera(const Path: string): Boolean;
  var
  SHFileOpStruct: TShFileOpStruct;
  DirBuf: array [0 .. 255] of Char;
  begin
  if DirectoryExists(Path) then
  try
  FillChar(SHFileOpStruct, Sizeof(SHFileOpStruct), 0);
  FillChar(DirBuf, Sizeof(DirBuf), 0);
  StrPCopy(DirBuf, Path);
  with SHFileOpStruct do
  begin
  Wnd := 0;
  pFrom := @DirBuf;
  wFunc := FO_DELETE;
  fFlags := FOF_NOCONFIRMATION or FOF_SILENT;
  end;
  Result := ShFileOperation(SHFileOpStruct) = 0;
  except
  Result := False;
  end;
  end; }
function ShF_ALaPapelera(Fichero: string): Boolean;
var
  FileOp: TShFileOpStruct;
  Resultado: Boolean;
begin
  Resultado := False;
  try
    try
      FillChar(FileOp, SizeOf(FileOp), #0);
      with FileOp do
      begin
        Wnd := Application.Handle;
        wFunc := FO_DELETE;
        pFrom := PChar(Fichero + #0#0);
        fFlags := FOF_SILENT or FOF_ALLOWUNDO or FOF_NOCONFIRMATION;
      end;
      Resultado := (ShFileOperation(FileOp) = 0);
    except
      on E: Exception do
        ShF_ControlDeError('[ShF.ShF_ALaPapelera][' + E.Message + ']');
    end;
  Finally
    Result := Resultado;
  end;
end;

procedure ShF_MaxPriori;
begin
  try
    SetPriorityClass(GetCurrentProcess, REALTIME_PRIORITY_CLASS);
    SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_TIME_CRITICAL);
  except
    ShowMessage('[E.ShF.'+glbNameEnc+'] Máxima prioridad del CPU.');
  end;
end;

procedure ShF_CentrarForm(Screen: TScreen);
var
  i: Integer;
begin
  try
    With Screen do
      for i := 0 to FormCount - 1 do
      Begin
        Forms[i].Top := Trunc((Height / 2) - (Forms[0].Height / 2));
        Forms[i].Left := Trunc((Width / 2) - (Forms[0].Width / 2));
      end;
  except
    ShowMessage('[E.ShF.ENUT09] Máxima prioridad del CPU.');
  end;
end;

function ShF_ListaArchivosTXT(directorioPadre: string): TStringList;
{ Saca todos los nombres de los archivos zip en un directorio }
var
  sr: TSearchRec;
begin
  Result := TStringList.Create;
  if FindFirst(directorioPadre + '\*', FaAnyfile, sr) = 0 then
    repeat
      if (sr.Attr and faDirectory = 0) or (sr.Name <> '.') and
        (sr.Name <> '..') then
      begin
        if UpperCase(AnsiRightStr(sr.Name, Length(sr.Name) - PosEx('.',
              sr.Name))) = 'TXT' then
        begin { Si es *.txt }
          Result.Add(AnsiLeftStr(sr.Name, PosEx('.', sr.Name) - 1));
          { Agregarlo en el TStringList }
        end;
      end;
    until FindNext(sr) <> 0;
    FindClose(sr);
end;

function ShF_ListaArchivosName(directorioPadre, Name: string): TStringList;
{ Saca todos los nombres de los archivos zip en un directorio }
var
  sr: TSearchRec;
begin
  Result := TStringList.Create;
  if FindFirst(directorioPadre + '\*', FaAnyfile, sr) = 0 then
    repeat
      if (sr.Attr and faDirectory = 0) or (sr.Name <> '.') and
        (sr.Name <> '..') then
      begin
        // if ( UpperCase(Trim(AnsiLeftStr(sr.Name, PosEx('.', sr.Name) -1 ))) = UpperCase(Trim(Name))) then begin
        if ShF_FindInStr(UpperCase(Trim(Name)),
          UpperCase(Trim(AnsiLeftStr(sr.Name, PosEx('.', sr.Name) - 1))),
          False) then
        begin
          Result.Add(sr.Name);
        end;
      end;
    until FindNext(sr) <> 0;
    FindClose(sr);
end;

function ShF_GenerarNombreConFecha(pref: String): String;
var
  Date, dd, mm, aaaa, hh, min, ss: String;
begin
  try
    Date := FormatDateTime('ddmmyyyy_HHmmss', Now());
    dd := AnsiLeftStr(Date, 2);
    mm := AnsiMidStr(Date, 3, 2);
    aaaa := AnsiMidStr(Date, 5, 4);
    hh := AnsiMidStr(Date, 10, 2);
    min := AnsiMidStr(Date, 12, 2);
    ss := AnsiMidStr(Date, 14, 2);
    Result := pref + aaaa + mm + dd + hh + min;
  except
    ShowMessage('[E.ShF.Nombre]');
  end
end;

Function ShF_EnviarEmail(Para, CC, Titulo: String; Cuerpo: TCaption;
  ConAdjunto: Boolean = False; FileAdjunto: String = ''): Boolean;
const
  olMailItem = 0;
var
  Outlook: OleVariant;
  vMailItem: variant;
  Band: Boolean;
begin
  try
    Outlook := GetActiveOleObject('Outlook.Application');
  except
    Outlook := CreateOleObject('Outlook.Application');
  end;
  vMailItem := Outlook.CreateItem(olMailItem);
  Para := Trim(Para);
  if AnsiRightStr(Para, 1) <> ';' then
    Para := Para + ';';
  Band := True;
  while Band do
  begin
    vMailItem.Recipients.Add(Trim(AnsiLeftStr(Para, PosEx(';', Para) - 1)));
    Para := AnsiRightStr(Para, Length(Para) - PosEx(';', Para));
    if Length(Para) = 0 then
      Band := False;
  end;
  vMailItem.CC := CC;
  vMailItem.Subject := Titulo;
  vMailItem.Body := Cuerpo;
  if ConAdjunto then
  begin
    FileAdjunto := Trim(FileAdjunto);
    if AnsiRightStr(FileAdjunto, 1) <> ';' then
      FileAdjunto := FileAdjunto + ';';
    Band := True;
    while Band do
    begin
      vMailItem.Attachments.Add(Trim(AnsiLeftStr(FileAdjunto, PosEx(';',
              FileAdjunto) - 1)));
      FileAdjunto := AnsiRightStr(FileAdjunto, Length(FileAdjunto) - PosEx(';',
          FileAdjunto));
      if Length(FileAdjunto) = 0 then
        Band := False;
    end;
  end;
  vMailItem.Send;
  VarClear(Outlook);
  VarClear(vMailItem);
end;

function ShF_IsNum(Cadena: string): Boolean;
var
  Cont: Integer;
  Num: Boolean;
  car: Char;
begin
  Cont := 1;
  Num := True;
  while Cont <= Length(Cadena) do
  begin
    car := Cadena[Cont];
    if StrScan('0123456789', car) = nil then
    begin
      Num := False;
    end;
    Inc(Cont);
  end;
  Result := Num;
end;

function ShF_ListaArchivosExt(directorioPadre, Ext: string): TStringList;
{ Saca todos los nombres de los archivos zip en un directorio }
var
  sr: TSearchRec;
begin
  Result := TStringList.Create;
  if FindFirst(directorioPadre + '\*', FaAnyfile, sr) = 0 then
    repeat
      if (sr.Attr and faDirectory = 0) or (sr.Name <> '.') and
        (sr.Name <> '..') then
      begin
        if UpperCase(AnsiRightStr(sr.Name, Length(sr.Name) - PosEx('.',
              sr.Name))) = UpperCase(Ext) then
        begin { Si es *.zip }
          Result.Add(AnsiLeftStr(sr.Name, PosEx('.', sr.Name) - 1));
          { Agregarlo en el TStringList }
        end;
      end;
    until FindNext(sr) <> 0;
    FindClose(sr);
end;

procedure ShF_File(Archivo, DirDestino, Accion: String);
var
  lpFileOp: TShFileOpStruct;
begin
  { Relleno de la estructura }
  // lpFileOp.Wnd := Self.Handle;
  if UpperCase(Accion) = 'COPIAR' then
    lpFileOp.wFunc := FO_COPY
  else if UpperCase(Accion) = 'MOVER' then
    lpFileOp.wFunc := FO_MOVE
  else if UpperCase(Accion) = 'BORRAR' then
    lpFileOp.wFunc := FO_DELETE
  else if UpperCase(Accion) = 'RENAME' then
    lpFileOp.wFunc := FO_RENAME
  else
    lpFileOp.wFunc := FO_COPY;
  lpFileOp.pFrom := PChar(Archivo + #0#0);
  lpFileOp.pTo := PChar(DirDestino + #0#0);
  lpFileOp.fFlags := FOF_SIMPLEPROGRESS or FOF_FILESONLY;
  lpFileOp.fAnyOperationsAborted := False;
  lpFileOp.hNameMappings := nil;
  // lpFileOp.lpszProgressTitle := PChar('Trasladando archivos al disco D' + #0#0);
  { Mover el archivo }
  ShFileOperation(lpFileOp);
end;

function ShF_GetUserName: String;
var
  pcUser: PChar;
  dwUSize: DWORD;
begin
  dwUSize := 21;
  GetMem(pcUser, dwUSize);
  try
    if Windows.GetUserName(pcUser, dwUSize) then
      Result := pcUser finally FreeMem(pcUser);
  end;
end;

function ShF_ListaArchivosNameExt(directorioPadre, Ext, Name: string)
  : TStringList;
var
  sr: TSearchRec;
begin
  Result := TStringList.Create;
  if FindFirst(directorioPadre + '\*', FaAnyfile, sr) = 0 then
    repeat
      if (sr.Attr and faDirectory = 0) or (sr.Name <> '.') and
        (sr.Name <> '..') then
      begin
        if UpperCase(AnsiRightStr(sr.Name, Length(sr.Name) - PosEx('.',
              sr.Name))) = UpperCase(Ext) then
        begin
          if ShF_FindInStr(UpperCase(Trim(Name)),
            UpperCase(Trim(AnsiLeftStr(sr.Name, PosEx('.', sr.Name) - 1))),
            False) then
          begin
            Result.Add(AnsiLeftStr(sr.Name, PosEx('.', sr.Name) - 1));
          end;
        end;
      end;
    until FindNext(sr) <> 0;
    FindClose(sr);
end;

function ShF_FindFile(const filespec: TFileName;
  attributes: Integer = 32): TStringList;
var
  spec: string;
  list: TStringList;

  procedure RFindFile(const folder: TFileName);
  var
    SearchRec: TSearchRec;
  begin
    // Busca todos los archivos concordantes
    // en la carpeta actual y agrega sus nombres
    // a la lista
    if FindFirst(folder + spec, attributes, SearchRec) = 0 then
    begin
      try
        repeat
          if (SearchRec.Attr and faDirectory = 0) or (SearchRec.Name <> '.')
            and (SearchRec.Name <> '..') then
            list.Add(folder + SearchRec.Name);
        until FindNext(SearchRec) <> 0;
      except
        FindClose(SearchRec);
        raise ;
      end;
      FindClose(SearchRec);
    end;
    // Ahora busca en las  subcarpetas
    if FindFirst(folder + '*', attributes Or faDirectory, SearchRec) = 0 then
    begin
      try
        repeat
          if ((SearchRec.Attr and faDirectory) <> 0) and
            (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
            RFindFile(folder + SearchRec.Name + '\');
        until FindNext(SearchRec) <> 0;
      except
        FindClose(SearchRec);
        raise ;
      end;
      FindClose(SearchRec);
    end;
  end;

// procedure RFindFile dentro de FindFile
begin // function FindFile
  list := TStringList.Create;
  try
    spec := ExtractFileName(filespec);
    RFindFile(extractfilepath(filespec));
    Result := list;
  except
    list.Free;
    raise ;
  end;
end;

function ShF_GetFileDate(TheFileName: string): string;
var
  FHandle: Integer;
begin
  FHandle := FileOpen(TheFileName, 0);
  try
    Result := DateTimeToStr(FileDateToDateTime(FileGetDate(FHandle)));
  finally
    FileClose(FHandle);
  end;
end;

function ShF_Ceros(Vari: string; lng: Integer): string;
begin
  while Length(Vari) < lng do
    Vari := '0' + Vari;
  Result := Vari;
end;

function ShF_GetTempDirectory: String;
var
  tempFolder: array [0 .. MAX_PATH] of Char;
begin
  GetTempPath(MAX_PATH, @tempFolder);
  Result := StrPas(tempFolder);

end;

function ShF_Session_ON(Session: TSession; pref: string;
  PrivateDir_NetFileDir: string = ''): String;
var
  Name, Dir: string;
  Dirs: TStringList;
  Cont: Integer;
begin
  Name := pref + FormatDateTime('ddmmyyyy_HHmmss', Now());
  if Length(Trim(PrivateDir_NetFileDir)) = 0 then
    PrivateDir_NetFileDir := ShF_GetTempDirectory;
  if not DirectoryExists(PrivateDir_NetFileDir) then
    PrivateDir_NetFileDir := ShF_GetTempDirectory;
  Dirs := TStringList.Create;
  Dirs := ShF_ListaDirectorios(PrivateDir_NetFileDir, pref);
  for Cont := 0 to Dirs.Count - 1 do
  begin
    ShF_ALaPapelera(PrivateDir_NetFileDir + Dirs[Cont] + '\*.*');
    RemoveDirectory(PChar(PrivateDir_NetFileDir + Dirs[Cont]));
  end;
  Dirs.Free;
  PrivateDir_NetFileDir := PrivateDir_NetFileDir + Name;
  CreateDir(PrivateDir_NetFileDir);
  ShF_ALaPapelera(PrivateDir_NetFileDir + '\*.*');
  glbTSession := TSession.Create(nil);
  Session.NetFileDir := PrivateDir_NetFileDir;
  Session.PrivateDir := PrivateDir_NetFileDir;
  Session.SessionName := Name;
  glbBanSession := True;
  Result := Name;
end;

procedure ShF_Session_OFF(Session: TSession; SessionName: string;
  PrivateDir_NetFileDir: string = '');
var
  Name, Dir: string;
begin
  Name := SessionName;
  Session.Free;
  Application.Terminate;
  if Length(Trim(PrivateDir_NetFileDir)) = 0 then
    PrivateDir_NetFileDir := ShF_GetTempDirectory;
  if not DirectoryExists(PrivateDir_NetFileDir) then
    PrivateDir_NetFileDir := ShF_GetTempDirectory;
  PrivateDir_NetFileDir := PrivateDir_NetFileDir + Name;
  ShF_ALaPapelera(PrivateDir_NetFileDir + '\*.*');
  ShF_BorrarCarpeta(PrivateDir_NetFileDir + '\');
  glbBanSession := False;
end;

procedure ShF_WordSetBold(RichE: TJvRichEdit; word: string);
var
  Text: string;
  Pos: Integer;
begin
  Text := RichE.Lines.Text;
  repeat
    Pos := ansipos(word, Text);
    RichE.SelStart := Pos;
    RichE.SelLength := Length(word);
    RichE.SelAttributes.Style := [fsBold];
    ShowMessage(AnsiMidStr(RichE.Text, RichE.SelStart, RichE.SelLength));
    Text[Pos] := Chr(255);
    Pos := ansipos(word, Text);
  until (Pos = 0);
end;

procedure ShF_WordSetColor(RichE: TJvRichEdit; word: string; col: TColor);
var
  Text: string;
  Pos: Integer;
begin
  Text := RichE.Lines.Text;
  repeat
    Pos := ansipos(word, Text);
    RichE.SelStart := Pos - 1;
    RichE.SelLength := Length(word);
    RichE.SelAttributes.Color := col;
    Text[Pos] := Chr(255);
    Pos := ansipos(word, Text);
  until (Pos = 0);
end;

function ShF_KillTask(ExeFileName: string): Integer;
const
  PROCESS_TERMINATE = $0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  Result := 0;
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  while Integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase
          (ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) = UpperCase
          (ExeFileName))) then
      Result := Integer(TerminateProcess(OpenProcess(PROCESS_TERMINATE,
            BOOL(0), FProcessEntry32.th32ProcessID), 0));
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;

function ShF_ProcessExists(ExeFileName: string): Boolean;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  Result := False;
  while Integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase
          (ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) = UpperCase
          (ExeFileName))) then
    begin
      Result := True;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;

function ShF_WebIng(URL: PChar): Boolean;
var
  hNet, hUrl: Pointer;
  BytesRead: DWORD;
  Buffer: array [0 .. 64] of Char;
begin
  Result := False;
  BytesRead := 0;
  if InternetAttemptConnect(0) <> ERROR_SUCCESS then
    exit;
  hNet := InternetOpen('WebIng', INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
  if hNet <> nil then
  begin
    hUrl := InternetOpenUrl(hNet, URL, nil, 0, INTERNET_FLAG_RELOAD
      { or INTERNET_FLAG_NO_AUTH } , 0);
    if hUrl <> nil then
    begin
      ZeroMemory(@Buffer[0], SizeOf(Buffer));
      Result := InternetReadFile(hUrl, @Buffer[0], SizeOf(Buffer), BytesRead);
      Result := Result and (BytesRead > 0);
      InternetCloseHandle(hUrl);
    end;
    InternetCloseHandle(hNet);
  end;
  Result := Result and (Pos('Access Denied', Buffer) = 0);
end;

function ShF_FicheroEnUso(Fichero: string): Boolean;
var
  HFileRes: HFILE;
  Res: string[6];
begin
  Result := False;

  HFileRes := CreateFile(PChar(Fichero), GENERIC_READ or GENERIC_WRITE, 0, nil,
    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
  Result := (HFileRes = INVALID_HANDLE_VALUE);
  if not Result then
    CloseHandle(HFileRes);
end;

{ ************************************************
  ***************** SISTEMA CAPI ******************
  ************************************************* }

function ShF_CAPI_ListaINI(Rubro: string): TStringList;
var
  Cont: Integer;
  Campo, Valor: string;
begin
  Result := TStringList.Create;
  { for Cont := 1 to StrToInt(ShF_LeerINI(glbINIConfig, Rubro, 'Tope')) do
    begin
    Campo := IntToStr(Cont);
    while Length(Campo) < 3 do
    Campo := '0' + Campo;
    Result.Add(ShF_LeerINI(glbINIConfig, Rubro, Campo));
    end; }
  Valor := '';
  for Cont := 1 to 999 do
  begin
    Campo := ShF_Ceros(IntToStr(Cont), 3);
    Valor := ShF_LeerINI(glbINIConfig, Rubro, Campo);
    if Valor = '#NoCampo#' then
      Break
    else
      Result.Add(Valor);
  end;
end;

procedure ShF_CAPI_EstructuraUSB(Raiz: string);
begin
  try
    if Length(Trim(Raiz)) = 0 then
      exit;
    {
    if Not DirectoryExists(Raiz + ':\' + glbNameSys) then
      CreateDir(Raiz + ':\' + glbNameSys);
    if Not DirectoryExists(Raiz + ':\' + glbNameSys + glbConst_DirEnvioOC) then
      CreateDir(Raiz + ':\' + glbNameSys + glbConst_DirEnvioOC);
    if Not DirectoryExists
      (Raiz + ':\' + glbNameSys + glbConst_DirActualizacion) then
      CreateDir(Raiz + ':\' + glbNameSys + glbConst_DirActualizacion);
    if Not DirectoryExists
      (Raiz + ':\' +
        glbNameSys + glbConst_DirSincronizacion) then
      CreateDir(Raiz + ':\' + glbNameSys + glbConst_DirSincronizacion);
    WinExec(PansiChar('cmd ' + Raiz + ':>attrib -s -h -r ' + glbNameSys),
      SW_HIDE);
    WinExec(PansiChar('cmd ' + Raiz + ':>attrib -s -h -r autorun.inf'),
      SW_HIDE);
    WinExec(PansiChar('cmd ' + Raiz + ':>del autorun.inf'), SW_HIDE);
    if FileExists(Raiz + ':\autorun.inf') then
    begin
      FileSetAttr(Raiz + ':\autorun.inf', faArchive);
      deletefile(Raiz + ':\autorun.inf');
    end;
    CreateDir(Raiz + ':\autorun.inf');
    FileSetAttr(Raiz + ':\autorun.inf', faHidden);    }
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_CAPI_EstructuraUSB][' + E.Message + ']');
  end
end;

function ShF_DateTimeToStr(Fecha: TDateTime): String;
var
  dd, mm, yyyy, hh, min, ss, ms: word;
  Dia, Mes, Año, Hora, Minuto, Segundo, MSegundo: string;
begin
  try
    DecodeDateTime(Fecha, yyyy, mm, dd, hh, min, ss, ms);
    Dia := ShF_Ceros(IntToStr(dd), 2);
    Mes := ShF_Ceros(IntToStr(mm), 2);
    Año := ShF_Ceros(IntToStr(yyyy), 4);
    Hora := ShF_Ceros(IntToStr(hh), 2);
    Minuto := ShF_Ceros(IntToStr(min), 2);
    Result := Dia + Mes + Año + Hora + Minuto;
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_DateTimeToStr][' + E.Message + ']');
  end
end;

function ShF_StrToDateTime(Fecha: String): TDateTime;
var
  FResultado: TDateTime;
  Dia, Mes, Año, Hora, Minuto, Segundo, MSegundo: Integer;
begin
  try
    Dia := StrToInt(AnsiLeftStr(Fecha, 2));
    Mes := StrToInt(AnsiMidStr(Fecha, 3, 2));
    Año := StrToInt(AnsiMidStr(Fecha, 5, 4));
    Hora := StrToInt(AnsiMidStr(Fecha, 9, 2));
    Minuto := StrToInt(AnsiMidStr(Fecha, 11, 2));
    FResultado := EncodeDateTime(Año, Mes, Dia, Hora, Minuto, 0, 0);
    Result := FResultado;
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_StrToDateTime][' + E.Message + ']');
  end
end;

function ShF_CAPI_InsertarTMov(Fecha, FechaF, Tipo, Origen, Destino,
  FileName: String): Integer;
var
  Qry: TQuery;
  Id: Integer;
  SQL: String;
begin
  Id := 0;
  try
    Qry := TQuery.Create(nil);
    if ShF_BDE_OpenSQL(ShF_SQL(11), Qry) then
    begin
      if Length(Trim(Qry.Fields[0].AsString)) = 0 then
      begin
        Id := 1;
      end
      else
      begin
        Id := Qry.Fields[0].AsInteger + 1;
      end;
    end;
    ShF_BDE_CloseSQL(Qry);
    SQL := 'INSERT INTO "' + glbTabMov + '" VALUES(' + IntToStr(Id)
      + ',''' + Fecha + ''' ,''' + FechaF + ''',''' + Tipo + ''',''' +
      Origen + ''',''' + Destino + ''',''' + FileName + ''')';
    ShF_BDE_EjecutarSQL(SQL);
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_CAPI_InsertarTMov][' + E.Message + ']');
  end;
  Result := Id;
end;

function ShF_CAPI_ResulxVivienda(jTRes_entrevista, Control, Viv_Sel: string)
  : String;
var
  QryA: TQuery;
  Resul: String;
begin
  Resul := '';
  try
    QryA := TQuery.Create(nil);
    if ShF_BDE_OpenSQL(ShF_SQL(12, jTRes_entrevista, Control, Viv_Sel), QryA)
      then
    begin
      QryA.First;
      Resul := QryA.Fields[3].AsString;
    end;
    ShF_BDE_CloseSQL(QryA);
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_CAPI_ResulxVivienda][' + E.Message + ']');
  end;
  Result := Resul;
end;

procedure ShF_CAPI_InsertarTListMov(n_mov: Integer;
  Control, Viv_Sel, Resul_V, Accion: String);
var
  SQL: String;
begin
  try
    SQL := 'INSERT INTO "' + glbTabListMov + '" VALUES(''' + IntToStr(n_mov)
      + ''',''' + Control + ''',''' + Viv_Sel + ''',''' + Resul_V + ''',''' +
      Accion + ''','''')';
    ShF_BDE_EjecutarSQL(SQL);
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_CAPI_InsertarTListMov][' + E.Message + ']');
  end
end;

function ShF_CAPI_GeneraUVisita(TRes_entrevista, TablaDestino,
  VivHogPer: string): Boolean;
var
  Qry_Reg, Qry_Ultimo: TQuery;
  flag, CreaTabla: Boolean;
  Filtro, Campos, Where: String;
begin
  flag := False;
  CreaTabla := True;
  try
    Case VivHogPer[1] Of
      'V':
        begin
          Filtro := 'Resul_V <> ""';
          Campos := 'Control, Viv_Sel';
        end;
      'H':
        begin
          Filtro := 'Resul_H <> ""';
          Campos := 'Control, Viv_Sel, Hogar';
        end;
      'P':
        begin
          Filtro := 'Resul_P <> ""';
          Campos := 'Control, Viv_Sel, Hogar, N_Ren';
        end;
    else
      Filtro := 'Resul_V <> ""';
      Campos := 'Control, Viv_Sel';
    end;
    Qry_Reg := TQuery.Create(nil);
    if ShF_BDE_OpenSelect(Qry_Reg, Campos, TRes_entrevista, Filtro,
      'GROUP BY ' + Campos) then
    begin
      while not Qry_Reg.Eof do
      begin
        Where := ' Control = "' + Qry_Reg.FieldByName('Control').AsString + '"';
        Where := Where + ' AND Cast(Viv_Sel as int) = ' + Qry_Reg.FieldByName
          ('Viv_Sel').AsString;
        Case VivHogPer[1] Of
          'V':
            begin
              Where := Where + ' AND Resul_V <> ""';
            end;
          'H':
            begin
              Where := Where + ' AND Cast(Hogar as int) = ' +
                Qry_Reg.FieldByName('Hogar').AsString;
              Where := Where + ' AND Resul_H <> ""';
            end;
          'P':
            begin
              Where := Where + ' AND Cast(Hogar as int) = ' +
                Qry_Reg.FieldByName('Hogar').AsString;
              Where := Where + ' AND Cast(N_Ren as int) = ' +
                Qry_Reg.FieldByName('N_Ren').AsString;
              Where := Where + ' AND Resul_P <> ""';
            end;
        else
          Where := Where + ' AND Resul_V <> ""';
        end;
        Qry_Ultimo := TQuery.Create(nil);
        if ShF_BDE_OpenSelect(Qry_Ultimo, '*', TRes_entrevista, Where,
          'Order By num_vis desc, fecha desc, Hra_2 desc') then
        begin
          if CreaTabla then
          begin
            if FileExists(TablaDestino) then
            begin
              deletefile(TablaDestino);
            end;
            ShF_CrearTabla(TablaDestino, Qry_Ultimo, ttDBase);
            CreaTabla := False;
            flag := True;
          end;
          Qry_Ultimo.First;
          ShF_BDE_Insert_First(TablaDestino, Qry_Ultimo);
        end;
        ShF_BDE_CloseSQL(Qry_Ultimo);
        Qry_Reg.Next;
      end;
    end;
    ShF_BDE_CloseSQL(Qry_Reg);
  except
    ShowMessage('[E.ShF.CAPI] Genera Última Visita');
  end;
  Result := flag;
end;

function ShF_ConexionFTP(Host, User, Pass: String): Boolean;
var
  FTP: TIdFTP;
  con: Boolean;
begin
  try
    con := False;
    FTP := TIdFTP.Create(nil);
    FTP.Username := User;
    FTP.Password := Pass;
    FTP.Host := Host;
    try
      FTP.Connect;
    except
      con := False;
    end;
    con := FTP.Connected;
    FTP.Disconnect;
    FTP.Free;
  except
    con := False;
  end;
  Result := con;
end;

procedure ShF_CAPI_VerifFTP(Host, User, Pass, Dir: String);
var
  FTP: TIdFTP;
begin
  try
    FTP := TIdFTP.Create(nil);
    FTP.Username := User;
    FTP.Password := Pass;
    FTP.Host := Host;
    try
      FTP.Connect;
    except
      raise Exception.Create('[E.ShF.ENSI10] Connect Estructura FTP.');
    end;
    if Length(Trim(glbFTPHostSubDir))> 0 then begin
       FTP.ChangeDir(glbFTPHostSubDir);
    end;


    if FTP.SendCmd('CWD ' + Dir, -1) <> -1 then
    begin
      if not(FTP.LastCmdResult.Code = '250') then
        FTP.MakeDir(Dir)
      else
        FTP.ChangeDirUp;
    end;
    {if FTP.SendCmd('CWD ' + Dir + glbConst_DirSincronizacion, -1) <> -1 then
    begin
      if not(FTP.LastCmdResult.Code = '250') then
        FTP.MakeDir(Dir + glbConst_DirSincronizacion)
      else
      begin
        FTP.ChangeDirUp;
        FTP.ChangeDirUp;
      end;
    end;

    if FTP.SendCmd('CWD ' + Dir + glbConst_DirEnvioOC, -1) <> -1 then
    begin
      if not(FTP.LastCmdResult.Code = '250') then
        FTP.MakeDir(Dir + glbConst_DirEnvioOC)
      else
      begin
        FTP.ChangeDirUp;
        FTP.ChangeDirUp;
      end;
    end;       }
    FTP.Disconnect;
    FTP.Free;
  except
    on E: Exception do
      ShF_ControlDeError('[ShF.ShF_CAPI_VerifFTP][' + E.Message + ']');
  end;

end;

function ShF_CAPI_DescripResulV(Resul_V: String): String;
var
  Resul: String;
begin
  Resul := '';
  try
    if Length(Trim(Resul_V)) = 0 then
    begin
      Resul_V := '0';
    end;
    Case StrToInt(Resul_V) Of
      1:
        Resul := '01 Entrevista completa con victimización';
      2:
        Resul := '02 Entrevista completa sin victimización';
      3:
        Resul := '03 Entrevista sin información de la persona elegida';
      4:
        Resul := '04 Entrevista incompleta';
      5:
        Resul := '05 Vivienda con algún hogar pendiente';
      6:
        Resul := '06 Entrevista aplazada';
      7:
        Resul := '07 Informante inadecuado';
      8:
        Resul := '08 Ausencia de ocupantes';
      9:
        Resul := '09 Negativa';
      10:
        Resul := '10 Vivienda deshabitada';
      11:
        Resul := '11 Vivienda de uso temporal';
      12:
        Resul := '12 No existe la vivienda';
      13:
        Resul := '13 Área insegura';
      14:
        Resul := '14 Otra situación';
    else
      Resul := '';
    end;
  except
    ShowMessage('[E.ShF.ENSI10] Descripción del resultado de la vivienda.')
  end;
  Result := Resul;
end;

function ShF_CAPI_AccionJE(Accion: String): String;
var
  Resul: String;
begin
  Resul := '';
  try
    if Length(Trim(Accion)) = 0 then
    begin
      Accion := '0';
    end;
    Case StrToInt(Accion) Of
      1:
        Resul := '1 Confirmar Resultado';
      2:
        Resul := '2 Mod. Resultado pero sigue NE';
      3:
        Resul := '3 Recupera Supervisor';
      4:
        Resul := '4 Recupera mismo Entrevistador';
      5:
        Resul := '5 Recupera otro Entrevistador';
      6:
        Resul := '6 No se Verificó';
    else
      Resul := '';
    end;
  except
    ShowMessage('[E.ShF.CAPI] Descripción de acción de VNR.');
  end;
  Result := Resul;
end;

function ShF_CAPI_FechaUVisita(Usuario, Control, Viv_Sel: string): String;
var
  QryA: TQuery;
  Resul: String;
begin
  Resul := '';
  try
    QryA := TQuery.Create(nil);
    if ShF_BDE_OpenSQL(ShF_SQL(12,
        glbDirUsu + '\' + Usuario + '\' + glbTabResEntrevista, Control,
        Viv_Sel), QryA) then
    begin
      QryA.First;
      Resul := QryA.Fields[1].AsString;
    end;
    ShF_BDE_CloseSQL(QryA);
  except
    ShowMessage('[E.ShF.CAPI] Descripción de acción de VNR.');
  end;
  Result := Resul;
end;

function ShF_CAPI_Num_Vis_UVisita(Usuario, Control, Viv_Sel: string): String;
var
  QryA: TQuery;
  Resul: String;
begin
  Resul := '';
  try
    QryA := TQuery.Create(nil);
    if ShF_BDE_OpenSQL(ShF_SQL(12,
        glbDirUsu + '\' + Usuario + '\' + glbTabResEntrevista, Control,
        Viv_Sel), QryA) then
    begin
      QryA.First;
      Resul := QryA.Fields[0].AsString;
    end;
    ShF_BDE_CloseSQL(QryA);
  except
    ShowMessage('[E.ShF.CAPI] Descripción de acción de VNR.');
  end;
  Result := Resul;
end;

function ShF_CAPI_getUACT(Usuario: string): string;
var
  Qry: TQuery;
  Rst: string;
begin
  Rst := '';
  Qry := TQuery.Create(nil);
  if ShF_BDE_OpenSQL(ShF_SQL(17, Usuario), Qry) then
  begin
    Qry.First;
    Rst := Qry.Fields[0].AsString;
  end;
  ShF_BDE_CloseSQL(Qry);
  Result := Rst;
end;

procedure ShF_CAPI_setUACT(Usuario, Date: string);
var
  SQL: string;
begin
  SQL := 'UPDATE "' + glbTabUsuario + '" SET UACT="' + Date +
    '" WHERE CVE_USU="' + Usuario + '"';
  ShF_BDE_EjecutarSQL(SQL);
end;


function ShF_CAPI_getVTab(Usuario: string): string;
var
  Qry: TQuery;
  Rst: string;
begin
  Rst := '';
  Qry := TQuery.Create(nil);
  if ShF_BDE_OpenSQL(ShF_SQL(22, Usuario), Qry) then
  begin
    Qry.First;
    Rst := Qry.Fields[0].AsString;
  end;
  ShF_BDE_CloseSQL(Qry);
  Result := Rst;
end;

procedure ShF_CAPI_setVTab(Usuario, Valor: string);
var
  SQL: string;
begin
  SQL := 'UPDATE "' + glbTabUsuario + '" SET V0_6="' + Valor +
    '" WHERE CVE_USU="' + Usuario + '"';
  ShF_BDE_EjecutarSQL(SQL);
end;


Function ShF_CAPI_ValidarVersionTablasUSU(DirUsu: string): Boolean;
var
  Res: Boolean;
  FIni, Campo, RutaTabla, Usuario: string;
  Tabla: TStringList;
  Cont, ContII: Integer;
  Qry: TQuery;
  ttab: TTable;
begin
  Res := True;
  try
    try
      if not DirectoryExists(DirUsu) then
      begin
        Res := False;
        exit;
      end;
      Tabla := TStringList.Create;
      Tabla.Clear;
      Campo := ShF_LeerINI(glbINIVer, 'VERSION', 'VersionT');
      Tabla := ShF_CAPI_ListaINI('Tabla_Usu');
      for Cont := 0 to Tabla.Count - 1 do
      begin
        if not Res then
          Break;
        RutaTabla := DirUsu + '\' + Tabla[Cont];
        if FileExists(RutaTabla) then
        begin
          ttab := TTable.Create(nil);
          ttab.TableName := RutaTabla;
          ttab.Open;
          if ttab.FindField(Campo) = nil then
          begin
            Res := False;
          end;
          ttab.Close;
          ttab.Free;
        end;
      end;
    except
      on E: Exception do
      begin
        ShF_ControlDeError
          ('[ShF.ShF_CAPI_ValidarVersionTablasUsu][' + E.Message + ']');
        Res := False;
      end;
    end;
  finally
    Tabla.Free;
    ShF_LiberarMemoria;
    Result := Res;
  end;
End;

end.
