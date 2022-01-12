unit frmSync;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, System.IniFiles,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.StdCtrls, Vcl.WinXCtrls, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MSSQL,
  FireDAC.Phys.MSSQLDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, System.UITypes;

type
  TformSyncService = class(TForm)
    Panel1: TPanel;
    Panel3: TPanel;
    StatusBar1: TStatusBar;
    Memo1: TMemo;
    Timer1: TTimer;
    ToggleSwitch: TToggleSwitch;
    FDConnet: TFDConnection;
    Panel2: TPanel;
    lbHours: TLabel;
    TrackBar1: TTrackBar;
    btnPlay: TSpeedButton;
    procedure Timer1Timer(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure btnPlayClick(Sender: TObject);
    procedure ToggleSwitchClick(Sender: TObject);
    procedure FDConnetBeforeConnect(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TrackBar1Change(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure SyncAjustesInventarios;
    procedure SyncTransferenciasEntreTiendas;
    procedure SyncIngresoTransferencia;
    procedure SyncMovimientoVentas;
    procedure SyncMovimientoSolicitudResultido;
    procedure SyncMovimientoTransferenciasCentralTienda;

    procedure SetInterval(Inteval: Integer);
    function GetInterval: Integer;
  public
    { Public declarations }
  end;

const
  Hour: Integer = 3600000;
  ConnStr: string =
    'DriverID=MSSQL;User_Name=%s;Database=%s;Server=%s;Password=%s;ApplicationName=Gazella Sync Services[Sync Mode]';

var
  formSyncService: TformSyncService;
  FIntervalToSync: Integer;
  OnExecute: Boolean;

implementation

{$R *.dfm}

procedure TformSyncService.btnPlayClick(Sender: TObject);
begin
  Timer1Timer(nil);
end;

procedure TformSyncService.btnStopClick(Sender: TObject);
begin
  Timer1.Enabled := false;
end;

procedure TformSyncService.FDConnetBeforeConnect(Sender: TObject);
begin
  //
  FDConnet.Params.LoadFromFile(ExtractFilePath(Application.ExeName) +
    'SyncInterface.connection');
end;

procedure TformSyncService.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if MessageDlg('Confirme que desea salie de la aplicacion.  Salir Ahora?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrNo then
    abort
  else
  begin
    MessageDlg('Se ha detenido la sincronizacion de interfaces.', mtWarning,
      [mbOk], 0, mbOk);
  end
end;

procedure TformSyncService.FormShow(Sender: TObject);
begin
  // Set interval
  try
    TrackBar1.Position := trunc(GetInterval / Hour);
    TrackBar1Change(nil)
  except
    TrackBar1.Position := 1;
  end;

  FDConnet.Close;
  FDConnet.Connected := true;
  StatusBar1.Panels[1].Text := 'Conectado con: ' + FDConnet.Params.Database;
end;

function TformSyncService.GetInterval: Integer;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.INI'));
  try
    FIntervalToSync := Ini.ReadInteger('CONFIG', 'SyncIntervalo', 0);
  finally
    Ini.Free;
    Result := FIntervalToSync;
  end;
end;

procedure TformSyncService.SetInterval(Inteval: Integer);
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.INI'));
  try
    Ini.WriteInteger('CONFIG', 'SyncIntervalo', Inteval);
  finally
    Ini.Free;
    FIntervalToSync := Inteval;
    Timer1.Enabled := false;
    Timer1.Interval := FIntervalToSync;
    Timer1.Enabled := true;
  end;
end;

procedure TformSyncService.SyncAjustesInventarios;
var
  spc_StoreProcedure: TFDStoredProc;
begin
  // Execute procedure spc_Sync_Ajustes_Inventarios
  spc_StoreProcedure := TFDStoredProc.Create(nil);
  try
    with spc_StoreProcedure do
    begin
      connection := FDConnet;
      SchemaName := 'dbo';
      StoredProcName := 'spc_Sync_Ajustes_Inventarios';
      Close;
      Prepared := true;
      Open;
      Memo1.Lines.Add('Proceso ID: 4 - Movimientos de Ajuste de Inventarios ');
      Memo1.Lines.Add('Inicio Actualizacion: ' + DateTimeToStr(now));
      Memo1.Lines.Add('Tabla: Int_Transacciones  ' +
        FieldByName('Int_Transacciones_Count').AsString +
        '  Registros Actualizados');
      Memo1.Lines.Add('Tabla: Int_Detalle_Transacciones  ' +
        FieldByName('Int_Detalle_Transacciones_Count').AsString +
        '  Registros Actualizados');
    end;
  finally
    spc_StoreProcedure.Free;

    Memo1.Lines.Add('Fin de Actualizacion: ' + DateTimeToStr(now));
    Memo1.Lines.Add
      ('-----------------------------------------------------------');
    Memo1.Lines.Add('');
  end;
end;

procedure TformSyncService.SyncIngresoTransferencia;
var
  spc_StoreProcedure: TFDStoredProc;
begin
  // Execute procedure  spc_Sync_Movimiento_Ventas
  spc_StoreProcedure := TFDStoredProc.Create(nil);
  try
    with spc_StoreProcedure do
    begin
      connection := FDConnet;
      SchemaName := 'dbo';
      StoredProcName := 'spc_Sync_Movimiento_Ingreso_Transferencia_Tiendas';
      Close;
      Prepared := true;
      Open;
      Memo1.Lines.Add
        ('Proceso ID: 3 - Traslados Ingresos de producto TIENDA – TIENDA');
      Memo1.Lines.Add('Inicio Actualizacion: ' + DateTimeToStr(now));
      Memo1.Lines.Add('Tabla: Int_Movimientos  ' +
        FieldByName('Int_Movimientos_Count').AsString +
        '  Registros Actualizados');
      Memo1.Lines.Add('Tabla: Int_Detalle_Movimientos  ' +
        FieldByName('Int_Detalle_Movimientos_Count').AsString +
        '  Registros Actualizados');
    end;
  finally
    spc_StoreProcedure.Free;

    Memo1.Lines.Add('Fin de Actualizacion: ' + DateTimeToStr(now));
    Memo1.Lines.Add
      ('-----------------------------------------------------------');
    Memo1.Lines.Add('');
  end;

end;

procedure TformSyncService.SyncMovimientoSolicitudResultido;
var
  spc_StoreProcedure: TFDStoredProc;
begin
  // Execute procedure  spc_Sync_Movimiento_Ventas
  spc_StoreProcedure := TFDStoredProc.Create(nil);
  try
    with spc_StoreProcedure do
    begin
      connection := FDConnet;
      SchemaName := 'dbo';
      StoredProcName := 'spc_Sync_Movimiento_Solicitud_Resultido';
      Close;
      Prepared := true;
      Open;
      Memo1.Lines.Add
        ('Proceso ID: 5 - Movimientos para solicitud de resultido tiendas');
      Memo1.Lines.Add('Inicio Actualizacion: ' + DateTimeToStr(now));
      Memo1.Lines.Add('Tabla: Int_Movimientos  ' +
        FieldByName('Int_Movimientos_Count').AsString +
        '  Registros Actualizados');
      Memo1.Lines.Add('Tabla: Int_Detalle_Movimientos  ' +
        FieldByName('Int_Detalle_Movimientos_Count').AsString +
        '  Registros Actualizados');
    end;
  finally
    spc_StoreProcedure.Free;

    Memo1.Lines.Add('Fin de Actualizacion: ' + DateTimeToStr(now));
    Memo1.Lines.Add
      ('-----------------------------------------------------------');
    Memo1.Lines.Add('');
  end;

end;

procedure TformSyncService.SyncMovimientoTransferenciasCentralTienda;
var
  spc_StoreProcedure: TFDStoredProc;
begin
  // Execute procedure  spc_Sync_Movimiento_Ventas
  spc_StoreProcedure := TFDStoredProc.Create(nil);
  try
    with spc_StoreProcedure do
    begin
      connection := FDConnet;
      SchemaName := 'dbo';
      StoredProcName := 'spc_Sync_Movimiento_Transferencia_Central_Tiendas';
      Close;
      Prepared := true;
      Open;
      Memo1.Lines.Add
        ('Proceso ID: 6 - Movimientos de Transferencias Central -> Tiendas');
      Memo1.Lines.Add('Inicio Actualizacion: ' + DateTimeToStr(now));
      Memo1.Lines.Add('Tabla: Int_Movimientos  ' +
        FieldByName('Int_Movimientos_Count').AsString +
        '  Registros Actualizados');
      Memo1.Lines.Add('Tabla: Int_Detalle_Movimientos  ' +
        FieldByName('Int_Detalle_Movimientos_Count').AsString +
        '  Registros Actualizados');
    end;
  finally
    spc_StoreProcedure.Free;

    Memo1.Lines.Add('Fin de Actualizacion: ' + DateTimeToStr(now));
    Memo1.Lines.Add
      ('-----------------------------------------------------------');
    Memo1.Lines.Add('');
  end;

end;

procedure TformSyncService.SyncMovimientoVentas;
var
  spc_StoreProcedure: TFDStoredProc;
begin
  // Execute procedure  spc_Sync_Movimiento_Ventas
  spc_StoreProcedure := TFDStoredProc.Create(nil);
  try
    with spc_StoreProcedure do
    begin
      connection := FDConnet;
      SchemaName := 'dbo';
      StoredProcName := 'spc_Sync_Movimiento_Ventas';
      Close;
      Prepared := true;
      Open;
      Memo1.Lines.Add
        ('Proceso ID: 1 - Movimientos de Ventas y Notas de Creditos');
      Memo1.Lines.Add('Inicio Actualizacion: ' + DateTimeToStr(now));
      Memo1.Lines.Add('Tabla: Int_Movimientos  ' +
        FieldByName('Int_Movimientos_Count').AsString +
        '  Registros Actualizados');
      Memo1.Lines.Add('Tabla: Int_Detalle_Movimientos  ' +
        FieldByName('Int_Detalle_Movimientos_Count').AsString +
        '  Registros Actualizados');
      Memo1.Lines.Add('Tabla: Int_Forma_Pagos  ' +
        FieldByName('Int_Forma_Pagos_Count').AsString +
        '  Registros Actualizados');
    end;
  finally
    spc_StoreProcedure.Free;

    Memo1.Lines.Add('Fin de Actualizacion: ' + DateTimeToStr(now));
    Memo1.Lines.Add
      ('-----------------------------------------------------------');
    Memo1.Lines.Add('');
  end;

end;

procedure TformSyncService.SyncTransferenciasEntreTiendas;
var
  spc_StoreProcedure: TFDStoredProc;
begin
  // Execute procedure  spc_Sync_Movimiento_Ventas
  spc_StoreProcedure := TFDStoredProc.Create(nil);
  try
    with spc_StoreProcedure do
    begin
      connection := FDConnet;
      SchemaName := 'dbo';
      StoredProcName := 'spc_Sync_Movimiento_Transferencia_Tiendas';
      Close;
      Prepared := true;
      Open;
      Memo1.Lines.Add
        ('Proceso ID: 2 - Traslados salida de producto TIENDA – TIENDA');
      Memo1.Lines.Add('Inicio Actualizacion: ' + DateTimeToStr(now));
      Memo1.Lines.Add('Tabla: Int_Movimientos  ' +
        FieldByName('Int_Movimientos_Count').AsString +
        '  Registros Actualizados');
      Memo1.Lines.Add('Tabla: Int_Detalle_Movimientos  ' +
        FieldByName('Int_Detalle_Movimientos_Count').AsString +
        '  Registros Actualizados');
    end;
  finally
    spc_StoreProcedure.Free;

    Memo1.Lines.Add('Fin de Actualizacion: ' + DateTimeToStr(now));
    Memo1.Lines.Add
      ('-----------------------------------------------------------');
    Memo1.Lines.Add('');
  end;

end;

procedure TformSyncService.Timer1Timer(Sender: TObject);
begin
  OnExecute := true;
  btnPlay.Enabled := not(OnExecute);
  Timer1.Enabled := not(OnExecute);

  Memo1.Lines.Add('****INICIO****');

  // Seccion #1   -> Doc Word
  SyncMovimientoVentas;

  // Seccion #2
  SyncTransferenciasEntreTiendas;

  // Seccion #3
  SyncIngresoTransferencia;

  // Seccion #4
  SyncAjustesInventarios;

  // Seccion #5
  SyncMovimientoSolicitudResultido;

  // Seccion #6
  SyncMovimientoTransferenciasCentralTienda;

  Memo1.Lines.Add('****FIN****');

  OnExecute := false;
  Timer1.Enabled := not(OnExecute);
  btnPlay.Enabled := not(OnExecute);
end;

procedure TformSyncService.ToggleSwitchClick(Sender: TObject);
begin
  Timer1.Enabled := not(ToggleSwitch.State in [tssOff]);
  btnPlay.Enabled := Timer1.Enabled;

  if (Timer1.Enabled) then
    StatusBar1.Panels[0].Text := 'Servicio Iniciado'
  else
    StatusBar1.Panels[0].Text := 'Servicio Detenido';
end;

procedure TformSyncService.TrackBar1Change(Sender: TObject);
begin
  SetInterval(TrackBar1.Position * Hour);
  lbHours.Caption := 'Sincronización cada  ' + TrackBar1.Position.ToString +
    '  Horas.';
end;

end.
