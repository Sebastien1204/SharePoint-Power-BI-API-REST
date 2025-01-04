unit Envoi_donnees;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, System.Generics.Collections, System.JSON, REST.Client, REST.Types, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, IdHTTP, IdSSLOpenSSL, IdGlobal;

type
  TForm1 = class(TForm)
    lblSelection : TLabel;
    cmbPlatform : TComboBox;
    grpIdentification : TGroupBox;
    lblTenantID : TLabel;
    edtTenantID : TEdit;
    lblClientID : TLabel;
    edtClientID : TEdit;
    lblClientSecret : TLabel;
    edtClientSecret : TEdit;
    grpSharePoint : TGroupBox;
    lblSiteID : TLabel;
    edtSiteID : TEdit;
    lblDriveID : TLabel;
    edtDriveID : TEdit;
    lblFilePath : TLabel;
    edtFilePath : TEdit;
    grpPowerBI : TGroupBox;
    lblWorkspaceID : TLabel;
    edtWorkspaceID : TEdit;
    lblDatasetID : TLabel;
    edtDatasetID : TEdit;
    lblDatasetAAjouter: TLabel;
    edtDatasetAAjouter: TMemo;
    btnGetToken : TButton;
    btnListSites : TButton;
    btnListDirectories : TButton;
    btnListFiles : TButton;
    btnUploadFile : TButton;
    btnListWorkspaces : TButton;
    btnListDatasetsAndReports: TButton;
    btnDeleteDataset : TButton;
    btnAddData : TButton;
    MemoOutput : TMemo;
    procedure cmbPlatformChange(Sender : TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnGetTokenClick(Sender : TObject);
    procedure btnListSitesClick(Sender : TObject);
    procedure btnListDirectoriesClick(Sender : TObject);
    procedure btnListFilesClick(Sender : TObject);
    procedure btnUploadFileClick(Sender : TObject);
    procedure btnListWorkspacesClick(Sender : TObject);
    procedure btnListDatasetsAndReportsClick(Sender : TObject);
    procedure btnDeleteDatasetClick(Sender : TObject);
    procedure btnAddDataClick(Sender : TObject);
  private
    AccessToken : string;
    Sites : string;
    Directories : string;
    Files : string;
    Workspaces : string;
    Datasets : string;
    procedure GetToken;
    procedure ListSites;
    procedure ListDirectories;
    procedure ListFiles;
    procedure UploadFile;
    procedure ListWorkspaces;
    procedure ListDatasetsAndReports;
    procedure DeleteDataset;
    procedure AddData;
  public
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.cmbPlatformChange(Sender: TObject);
var
  JSONContent : string;
begin
  if cmbPlatform.ItemIndex = 0 then // SharePoint s�lectionn�
  begin
    edtTenantID.Text := '587ee63e-67d5-495f-b8e4-3608ddfb72cf'; // Valeurs SharePoint
    edtClientID.Text := 'c9c3efb9-11f2-4583-adac-c0305a247a42';
    edtClientSecret.Text := 'TOe8Q~SjZp8oTsGAZfzPODRNVGlNAGBNRugXic4k';
    edtWorkspaceID.Text := '';
    edtDatasetID.Text := '';
    edtDatasetAAjouter.Text := '';
  end
  else if cmbPlatform.ItemIndex = 1 then // Power BI s�lectionn�
  begin
    edtTenantID.Text := '587ee63e-67d5-495f-b8e4-3608ddfb72cf'; // Valeurs Power BI
    edtClientID.Text := '8cfe895d-05eb-4eae-92a9-6601fd9e8a8f';
    edtClientSecret.Text := 'X_38Q~-euwKIec9fmkS7hgYSX4ec~DR_8CmKdaPm';

    JSONContent :=
                  '{' +
                    '"name": "MyNewDataset",' +
                    '"defaultMode": "Push",' +
                    '"tables": [' +
                      '{' +
                        '"name": "Table1",' +
                        '"columns": [' +
                          '{"name": "ID","dataType": "Int64"},' +
                          '{"name": "Name","dataType": "String"},' +
                          '{"name": "Timestamp","dataType": "DateTime"}' +
                        ']' +
                      '}' +
                    ']' +
                  '}';

    edtDatasetAAjouter.Text := JSONContent;
    edtSiteID.Text := '';
    edtDriveID.Text := '';
    edtFilePath.Text := '';
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  // Initialisation des valeurs par d�faut
  edtTenantID.Text := '587ee63e-67d5-495f-b8e4-3608ddfb72cf';
    edtClientID.Text := 'c9c3efb9-11f2-4583-adac-c0305a247a42';
    edtClientSecret.Text := 'TOe8Q~SjZp8oTsGAZfzPODRNVGlNAGBNRugXic4k';

  // Associer l'�v�nement au changement de s�lection
  cmbPlatform.OnChange := cmbPlatformChange;

  // S�lectionner SharePoint par d�faut
  cmbPlatform.ItemIndex := 0;
  cmbPlatformChange(cmbPlatform);
end;

procedure TForm1.btnGetTokenClick(Sender : TObject);
begin
  MemoOutput.Clear;
  GetToken;
end;

procedure TForm1.GetToken;
var
  RESTClient : TRESTClient;
  RESTRequest : TRESTRequest;
  RESTResponse : TRESTResponse;
  JSONResponse : TJSONObject;
  TokenEndpoint, Scope : string;
begin

  TokenEndpoint := Format('https://login.microsoftonline.com/%s/oauth2/v2.0/token', [edtTenantID.Text]);

  // V�rifier la plateforme s�lectionn�e et d�finir le scope
  if cmbPlatform.ItemIndex = 0 then // SharePoint
  begin
    Scope := 'https://graph.microsoft.com/.default';
  end
  else if cmbPlatform.ItemIndex = 1 then // Power BI
  begin
    Scope := 'https://analysis.windows.net/powerbi/api/.default';
  end
  else
  begin
    MemoOutput.Lines.Add('S�lectionnez une plateforme avant de r�cup�rer le token.');
    Exit;
  end;

  // Cr�er les objets REST
  RESTClient := TRESTClient.Create(TokenEndpoint);
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);

  try
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmPOST;

    // Ajouter le header Content-Type
    RESTRequest.Params.AddItem('Content-Type', 'application/x-www-form-urlencoded', pkHTTPHEADER, [poDoNotEncode]);

    // Ajouter les param�tres du corps de la requ�te
    RESTRequest.AddParameter('client_id', edtClientID.Text, pkGETorPOST);
    RESTRequest.AddParameter('client_secret', edtClientSecret.Text, pkGETorPOST);
    RESTRequest.AddParameter('grant_type', 'client_credentials', pkGETorPOST);
    RESTRequest.AddParameter('scope', Scope, pkGETorPOST);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // Traiter la r�ponse
    JSONResponse := TJSONObject.ParseJSONValue(RESTResponse.Content) as TJSONObject;
    try
      if Assigned(JSONResponse) and JSONResponse.TryGetValue('access_token', AccessToken) then
      begin
        MemoOutput.Lines.Add('Token r�cup�r� avec succ�s :');
        MemoOutput.Lines.Add(AccessToken);
      end
      else
        MemoOutput.Lines.Add('Erreur lors de la r�cup�ration du token.');
    finally
      JSONResponse.Free;
    end;
  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;

  // Lib�rer les ressources
  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;
end;

procedure TForm1.btnListSitesClick(Sender : TObject);
begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;
  MemoOutput.Clear;
  ListSites;
end;

procedure TForm1.ListSites;
var
  RESTClient: TRESTClient;
  RESTRequest: TRESTRequest;
  RESTResponse: TRESTResponse;
  JSONResponse, Value: TJSONObject;
  JSONArray: TJSONArray;
  i: Integer;
  Endpoint: string;
begin
  Endpoint := 'https://graph.microsoft.com/v1.0/sites';

  RESTClient := TRESTClient.Create(Endpoint);
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);
  try
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmGET;

    // Ajouter l'en-t�te Authorization
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // Traiter la r�ponse
    JSONResponse := TJSONObject.ParseJSONValue(RESTResponse.Content) as TJSONObject;
    try
      if Assigned(JSONResponse) then
      begin
        if JSONResponse.TryGetValue('value', JSONArray) then
        begin
          MemoOutput.Lines.Add('Liste des sites :');
          MemoOutput.Lines.Add('');
          for i := 0 to JSONArray.Count - 1 do
          begin
            // Pour chaque site, r�cup�rer le displayName et l'id
            Value := JSONArray.Items[i] as TJSONObject;
            if Value.TryGetValue('displayName', Sites) then
            begin
              MemoOutput.Lines.Add('Nom : ' + Sites);
            end;
            if Value.TryGetValue('id', Sites) then
            begin
              MemoOutput.Lines.Add('ID : ' + Sites);
              MemoOutput.Lines.Add('');
            end;
          end;
        end
        else
          MemoOutput.Lines.Add('Erreur : Pas de donn�es site dans la r�ponse.');
      end
      else
        MemoOutput.Lines.Add('Erreur lors de l''affichage des sites.');
    finally
      JSONResponse.Free;
    end;
  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;
  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;
end;

procedure TForm1.btnListDirectoriesClick(Sender: TObject);
begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;

  if Trim(edtSiteID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un Site ID valide.');
    Exit;
  end;

  MemoOutput.Clear;
  ListDirectories;
end;

procedure TForm1.ListDirectories;
var
  RESTClient: TRESTClient;
  RESTRequest: TRESTRequest;
  RESTResponse: TRESTResponse;
  JSONResponse, Value: TJSONObject;
  JSONArray: TJSONArray;
  i: Integer;
  DriveName, DriveID, Endpoint: string;
begin
  Endpoint := Format('https://graph.microsoft.com/v1.0/sites/%s/drives', [edtSiteID.Text]);

  RESTClient := TRESTClient.Create(Endpoint);
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);
  try
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmGET;

    // Ajouter l'en-t�te Authorization
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // Traiter la r�ponse
    JSONResponse := TJSONObject.ParseJSONValue(RESTResponse.Content) as TJSONObject;
    try
      if Assigned(JSONResponse) then
      begin
        if JSONResponse.TryGetValue('value', JSONArray) then
        begin
          MemoOutput.Lines.Add('Liste des r�pertoires :');
          MemoOutput.Lines.Add('');
          for i := 0 to JSONArray.Count - 1 do
          begin
            // R�cup�rer chaque drive
            Value := JSONArray.Items[i] as TJSONObject;

            // Obtenir le nom du drive
            if Value.TryGetValue('name', DriveName) then
              MemoOutput.Lines.Add('Nom : ' + DriveName);

            // Obtenir l'ID du drive
            if Value.TryGetValue('id', DriveID) then
              MemoOutput.Lines.Add('ID : ' + DriveID);

            MemoOutput.Lines.Add(''); // Ligne vide pour s�parer chaque r�pertoire
          end;
        end
        else
          MemoOutput.Lines.Add('Erreur : Pas de donn�es de r�pertoire dans la r�ponse.');
      end
      else
        MemoOutput.Lines.Add('Erreur lors de l''affichage des r�pertoires.');
    finally
      JSONResponse.Free;
    end;
  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;
  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;
end;

procedure TForm1.btnListFilesClick(Sender : TObject);

begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;

  if Trim(edtSiteID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un Site ID valide.');
    Exit;
  end;

  if Trim(edtDriveID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un R�pertoire ID valide.');
    Exit;
  end;

  MemoOutput.Clear;
  ListFiles;
end;

procedure TForm1.ListFiles;
var
  RESTClient: TRESTClient;
  RESTRequest: TRESTRequest;
  RESTResponse: TRESTResponse;
  JSONResponse: TJSONObject;
  JSONArray: TJSONArray;
  FileItem: TJSONObject;
  i: Integer;
  FileName: string;
begin
  if edtSiteID.Text = '' then
  begin
    MemoOutput.Lines.Add('Site ID est requis.');
    Exit;
  end;

  if edtDriveID.Text = '' then
  begin
    MemoOutput.Lines.Add('Drive ID est requis.');
    Exit;
  end;

  // Configurer le client REST avec l'URL de l'API Graph
  RESTClient := TRESTClient.Create(Format('https://graph.microsoft.com/v1.0/sites/%s/drives/%s/root/children',
    [edtSiteID.Text, edtDriveID.Text]));
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);
  try
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmGET;

    // Ajouter le header d'authentification
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // Traiter la r�ponse JSON
    JSONResponse := TJSONObject.ParseJSONValue(RESTResponse.Content) as TJSONObject;
    try
      if Assigned(JSONResponse) and JSONResponse.TryGetValue('value', JSONArray) then
      begin
        MemoOutput.Lines.Add('Contenu du r�pertoire :');
        for i := 0 to JSONArray.Count - 1 do
        begin
          FileItem := JSONArray.Items[i] as TJSONObject;

          // R�cup�rer le nom du fichier ou du dossier
          if FileItem.TryGetValue('name', FileName) then
          begin
            // V�rifier si c'est un fichier ou un dossier
            if FileItem.TryGetValue('folder', JSONResponse) then
              MemoOutput.Lines.Add(Format('Dossier : %s', [FileName]))
            else if FileItem.TryGetValue('file', JSONResponse) then
              MemoOutput.Lines.Add(Format('Fichier : %s', [FileName]));
          end;
        end;
      end
      else
      begin
        MemoOutput.Lines.Add('Aucun fichier ou dossier trouv�.');
      end;
    finally
      JSONResponse.Free;
    end;
  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;

  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;
end;

procedure TForm1.btnUploadFileClick(Sender: TObject);
begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;

  if Trim(edtSiteID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un Site ID valide.');
    Exit;
  end;

  if Trim(edtDriveID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un R�pertoire ID valide.');
    Exit;
  end;

  MemoOutput.Clear;
  UploadFile;
end;

procedure TForm1.UploadFile;
var
  RESTClient : TRESTClient;
  RESTRequest : TRESTRequest;
  RESTResponse : TRESTResponse;
  FileStream : TFileStream;
  Endpoint : string;
  FilePath, FileName : string;
begin

  FilePath := edtFilePath.Text;

  // Supprimer le premier et le dernier caract�re
  FilePath := Copy(FilePath, 2, Length(FilePath) - 2);

  FileName := ExtractFileName(FilePath);

  if not FileExists(FilePath) then
  begin
    MemoOutput.Lines.Add('Erreur : le fichier sp�cifi� est introuvable.');
    Exit;
  end;

  Endpoint := Format('https://graph.microsoft.com/v1.0/sites/%s/drives/%s/root:/%s:/content',[edtSiteID.Text, edtDriveID.Text, FileName]);

  RESTClient := TRESTClient.Create(Endpoint);
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);
  FileStream := TFileStream.Create(FilePath, fmOpenRead);

  try

    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmPUT;

    //Ajouter l'en-t�te Authorization
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    //Ajouter le flux du fichier comme body de la requ�te
    RESTRequest.AddBody(FileStream, ctAPPLICATION_OCTET_STREAM);

    //Ex�cuter la requ�te
    RESTRequest.Execute;

    //V�rifier la r�ussite de la requ�te
    if (RESTResponse.StatusCode = 200) or (RESTResponse.StatusCode = 201) then

    begin
      MemoOutput.Lines.Add(Format('Fichier envoy� avec succ�s : %s', [FileName]));
    end

    else

    begin
      MemoOutput.Lines.Add('Erreur lors de l''envoi du fichier :');
      MemoOutput.Lines.Add(Format('Code HTTP : %d - %s', [RESTResponse.StatusCode, RESTResponse.StatusText]));
      MemoOutput.Lines.Add('R�ponse : ' + RESTResponse.Content);
    end;

  except

    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);

  end;

  FileStream.Free;
  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;

end;

procedure TForm1.btnListWorkspacesClick(Sender : TObject);
begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;

  MemoOutput.Clear;
  ListWorkspaces;
end;

procedure TForm1.ListWorkspaces;
var
  RESTClient: TRESTClient;
  RESTRequest: TRESTRequest;
  RESTResponse: TRESTResponse;
  JSONResponse, Value: TJSONObject;
  JSONArray: TJSONArray;
  i: Integer;
  Endpoint: string;
begin
  Endpoint := 'https://api.powerbi.com/v1.0/myorg/groups';

  RESTClient := TRESTClient.Create(Endpoint);
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);
  try
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmGET;

    // Ajouter l'en-t�te Authorization
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // Traiter la r�ponse
    JSONResponse := TJSONObject.ParseJSONValue(RESTResponse.Content) as TJSONObject;
    try
      if Assigned(JSONResponse) then
      begin
        if JSONResponse.TryGetValue('value', JSONArray) then
        begin
          MemoOutput.Lines.Add('Liste des espaces de travail :');
          MemoOutput.Lines.Add('');
          for i := 0 to JSONArray.Count - 1 do
          begin
            // Pour chaque espace de travail, r�cup�rer le name et l'id
            Value := JSONArray.Items[i] as TJSONObject;
            if Value.TryGetValue('name', Workspaces) then
            begin
              MemoOutput.Lines.Add('Nom : ' + Workspaces);
            end;
            if Value.TryGetValue('id', Workspaces) then
            begin
              MemoOutput.Lines.Add('ID : ' + Workspaces);
              MemoOutput.Lines.Add('');
            end;
          end;
        end
        else
          MemoOutput.Lines.Add('Erreur : Pas de donn�es d''espace de travail dans la r�ponse.');
      end
      else
        MemoOutput.Lines.Add('Erreur lors de l''affichage des espaces de travail.');
    finally
      JSONResponse.Free;
    end;
  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;
  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;
end;

procedure TForm1.btnListDatasetsAndReportsClick(Sender: TObject);
begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;

  if Trim(edtWorkspaceID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un Espace de travail ID valide.');
    Exit;
  end;

  MemoOutput.Clear;
  ListDatasetsAndReports;
end;

procedure TForm1.ListDatasetsAndReports;
var
  RESTClient: TRESTClient;
  RESTRequest: TRESTRequest;
  RESTResponse: TRESTResponse;
  JSONResponse, Value: TJSONObject;
  JSONArray: TJSONArray;
  i: Integer;
  DatasetName, DatasetID, ReportName, ReportID: string;
  Endpoint: string;
begin
  // Initialiser les objets pour la premi�re requ�te (datasets)
  RESTClient := TRESTClient.Create('');
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);

  try
    // Lister les datasets
    Endpoint := Format('https://api.powerbi.com/v1.0/myorg/groups/%s/datasets', [edtWorkspaceID.Text]);

    RESTClient.BaseURL := Endpoint;
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmGET;

    // Ajouter l'en-t�te Authorization
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // Traiter la r�ponse pour les datasets
    JSONResponse := TJSONObject.ParseJSONValue(RESTResponse.Content) as TJSONObject;
    try
      if Assigned(JSONResponse) then
      begin
        if JSONResponse.TryGetValue('value', JSONArray) then
        begin
          MemoOutput.Lines.Add('Liste des datasets :');
          MemoOutput.Lines.Add('');
          for i := 0 to JSONArray.Count - 1 do
          begin
            // R�cup�rer chaque dataset
            Value := JSONArray.Items[i] as TJSONObject;

            // Obtenir le nom du dataset
            if Value.TryGetValue('name', DatasetName) then
              MemoOutput.Lines.Add('Nom du dataset : ' + DatasetName);

            // Obtenir l'ID du dataset
            if Value.TryGetValue('id', DatasetID) then
              MemoOutput.Lines.Add('ID du dataset : ' + DatasetID);

            MemoOutput.Lines.Add(''); // Ligne vide pour s�parer chaque dataset
          end;
        end
        else
          MemoOutput.Lines.Add('Erreur : Pas de donn�es de dataset dans la r�ponse.');
      end
      else
        MemoOutput.Lines.Add('Erreur lors de l''affichage des datasets.');
    finally
      JSONResponse.Free;
    end;

  finally
    // Lib�rer les ressources apr�s les datasets
    RESTResponse.Free;
    RESTRequest.Free;
    RESTClient.Free;

    // R�initialiser les objets pour la seconde requ�te (reports)
    RESTClient := TRESTClient.Create('');
    RESTRequest := TRESTRequest.Create(nil);
    RESTResponse := TRESTResponse.Create(nil);
  end;

  try
    // Lister les rapports (reports)
    Endpoint := Format('https://api.powerbi.com/v1.0/myorg/groups/%s/reports', [edtWorkspaceID.Text]);

    RESTClient.BaseURL := Endpoint;
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmGET;

    // Ajouter l'en-t�te Authorization
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    // Ex�cuter la requ�te pour les rapports
    RESTRequest.Execute;

    // Traiter la r�ponse pour les rapports
    JSONResponse := TJSONObject.ParseJSONValue(RESTResponse.Content) as TJSONObject;
    try
      if Assigned(JSONResponse) then
      begin
        if JSONResponse.TryGetValue('value', JSONArray) then
        begin
          MemoOutput.Lines.Add('Liste des rapports :');
          MemoOutput.Lines.Add('');
          for i := 0 to JSONArray.Count - 1 do
          begin
            // R�cup�rer chaque rapport
            Value := JSONArray.Items[i] as TJSONObject;

            // Obtenir le nom du rapport
            if Value.TryGetValue('name', ReportName) then
              MemoOutput.Lines.Add('Nom du rapport : ' + ReportName);

            // Obtenir l'ID du rapport
            if Value.TryGetValue('id', ReportID) then
              MemoOutput.Lines.Add('ID du rapport : ' + ReportID);

            MemoOutput.Lines.Add(''); // Ligne vide pour s�parer chaque report
          end;
        end
        else
          MemoOutput.Lines.Add('Erreur : Pas de donn�es de rapport dans la r�ponse.');
      end
      else
        MemoOutput.Lines.Add('Erreur lors de l''affichage des rapports.');
    finally
      JSONResponse.Free;
    end;

  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;

  // Lib�rer les ressources pour la deuxi�me requ�te (reports)
  RESTResponse.Free;
  RESTRequest.Free;
  RESTClient.Free;
end;

procedure TForm1.btnDeleteDatasetClick(Sender: TObject);
begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;

  if Trim(edtWorkspaceID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un Espace de travail ID valide.');
    Exit;
  end;

  if Trim(edtDatasetID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un Dataset ID valide.');
    Exit;
  end;

  MemoOutput.Clear;
  DeleteDataset;
end;

procedure TForm1.DeleteDataset;
var
  RESTClient: TRESTClient;
  RESTRequest: TRESTRequest;
  RESTResponse: TRESTResponse;
  Endpoint: string;
begin
  // Construire l'URL de l'endpoint avec l'ID du workspace et du jeu de donn�es
  Endpoint := Format('https://api.powerbi.com/v1.0/myorg/groups/%s/datasets/%s', [edtWorkspaceID.Text, edtDatasetID.Text]);

  RESTClient := TRESTClient.Create(Endpoint);
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);

  try
    // Configurer la requ�te
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmDELETE;

    // Ajouter le header Authorization avec le token d'acc�s
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // V�rifier la r�ponse
    if (RESTResponse.StatusCode = 200) or (RESTResponse.StatusCode = 204) then
    begin
      MemoOutput.Lines.Add('Le jeu de donn�es a �t� supprim� avec succ�s.');
    end
    else
    begin
      MemoOutput.Lines.Add('Erreur lors de la suppression du jeu de donn�es :');
      MemoOutput.Lines.Add(Format('Code HTTP : %d - %s', [RESTResponse.StatusCode, RESTResponse.StatusText]));
      MemoOutput.Lines.Add('R�ponse : ' + RESTResponse.Content);
    end;
  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;

  // Lib�rer les ressources
  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;
end;

procedure TForm1.btnAddDataClick(Sender: TObject);
begin
  if AccessToken.IsEmpty then
  begin
    ShowMessage('Veuillez d''abord r�cup�rer le token.');
    Exit;
  end;

  if Trim(edtWorkspaceID.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un Espace de travail ID valide.');
    Exit;
  end;

  if Trim(edtDatasetAAjouter.Text) = '' then
  begin
    ShowMessage('Veuillez entrer un jeu de donn�es valide.');
    Exit;
  end;

  MemoOutput.Clear;
  AddData;
end;

procedure TForm1.AddData;
var
  RESTClient : TRESTClient;
  RESTRequest : TRESTRequest;
  RESTResponse : TRESTResponse;
  JSONResponse : TJSONObject;
  Endpoint : string;
begin

  Endpoint := Format('https://api.powerbi.com/v1.0/myorg/groups/%s/datasets', [edtWorkspaceID.Text]);

  // Cr�er les objets REST
  RESTClient := TRESTClient.Create(Endpoint);
  RESTRequest := TRESTRequest.Create(nil);
  RESTResponse := TRESTResponse.Create(nil);

  try
    RESTRequest.Client := RESTClient;
    RESTRequest.Response := RESTResponse;
    RESTRequest.Method := rmPOST;

    // Ajouter les en-t�tes Authorization et Content-Type
    RESTRequest.Params.AddItem('Authorization', 'Bearer ' + AccessToken, pkHTTPHEADER, [poDoNotEncode]);
    RESTRequest.Params.AddItem('Content-Type', 'application/json', pkHTTPHEADER, [poDoNotEncode]);

    // Ajouter le dataset au corps de la requ�te
    RESTRequest.AddBody(edtDatasetAAjouter.Text, ctAPPLICATION_JSON);

    // Ex�cuter la requ�te
    RESTRequest.Execute;

    // V�rifier la r�ponse
    if RESTResponse.StatusCode = 201 then
    begin
      MemoOutput.Lines.Add('Dataset ajout� avec succ�s.');
    end
    else
    begin
      MemoOutput.Lines.Add('Erreur lors de l''ajout du dataset :');
      MemoOutput.Lines.Add(Format('Code HTTP : %d - %s', [RESTResponse.StatusCode, RESTResponse.StatusText]));
      MemoOutput.Lines.Add('R�ponse : ' + RESTResponse.Content);
    end;

  except
    on E: Exception do
      MemoOutput.Lines.Add('Erreur : ' + E.Message);
  end;

  // Lib�rer les ressources
  RESTClient.Free;
  RESTRequest.Free;
  RESTResponse.Free;
end;

end.
