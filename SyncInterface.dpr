program SyncInterface;

uses
  Vcl.Forms,
  frmSync in 'frmSync.pas' {formSyncService},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Carbon');
  Application.CreateForm(TformSyncService, formSyncService);
  Application.Run;
end.
