object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'SharePoint & Power BI API REST'
  ClientHeight = 599
  ClientWidth = 800
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  OnClick = cmbPlatformChange
  TextHeight = 13
  object lblSelection: TLabel
    Left = 180
    Top = 16
    Width = 113
    Height = 13
    Caption = 'Choisir la plateforme :'
  end
  object cmbPlatform: TComboBox
    Left = 300
    Top = 13
    Width = 200
    Height = 21
    Style = csDropDownList
    TabOrder = 0
    OnChange = cmbPlatformChange
    Items.Strings = (
      'SharePoint'
      'Power BI')
  end
  object grpIdentification: TGroupBox
    Left = 16
    Top = 45
    Width = 774
    Height = 80
    Caption = 'Identification'
    TabOrder = 1
    object lblTenantID: TLabel
      Left = 16
      Top = 24
      Width = 55
      Height = 13
      Caption = 'Tenant ID :'
    end
    object lblClientID: TLabel
      Left = 276
      Top = 24
      Width = 50
      Height = 13
      Caption = 'Client ID :'
    end
    object lblClientSecret: TLabel
      Left = 536
      Top = 24
      Width = 70
      Height = 13
      Caption = 'Client Secret :'
    end
    object edtTenantID: TEdit
      Left = 16
      Top = 40
      Width = 244
      Height = 21
      TabOrder = 0
    end
    object edtClientID: TEdit
      Left = 276
      Top = 40
      Width = 244
      Height = 21
      TabOrder = 1
    end
    object edtClientSecret: TEdit
      Left = 536
      Top = 40
      Width = 208
      Height = 21
      PasswordChar = '*'
      TabOrder = 2
    end
  end
  object grpSharePoint: TGroupBox
    Left = 16
    Top = 130
    Width = 376
    Height = 183
    Caption = 'SharePoint'
    TabOrder = 5
    object lblSiteID: TLabel
      Left = 16
      Top = 24
      Width = 39
      Height = 13
      Caption = 'Site ID :'
    end
    object lblDriveID: TLabel
      Left = 16
      Top = 80
      Width = 74
      Height = 13
      Caption = 'R'#233'pertoire ID :'
    end
    object lblFilePath: TLabel
      Left = 16
      Top = 130
      Width = 149
      Height = 13
      Caption = 'Chemin du fichier '#224' envoyer :'
    end
    object edtSiteID: TEdit
      Left = 16
      Top = 40
      Width = 320
      Height = 21
      TabOrder = 0
    end
    object edtDriveID: TEdit
      Left = 16
      Top = 100
      Width = 320
      Height = 21
      TabOrder = 1
    end
    object edtFilePath: TEdit
      Left = 16
      Top = 150
      Width = 320
      Height = 21
      TabOrder = 2
    end
  end
  object grpPowerBI: TGroupBox
    Left = 416
    Top = 130
    Width = 376
    Height = 220
    Caption = 'Power BI'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    object lblWorkspaceID: TLabel
      Left = 16
      Top = 24
      Width = 106
      Height = 13
      Caption = 'Espace de Travail ID :'
    end
    object lblDatasetID: TLabel
      Left = 16
      Top = 104
      Width = 59
      Height = 13
      Caption = 'Dataset ID :'
    end
    object lblDatasetAAjouter: TLabel
      Left = 16
      Top = 65
      Width = 130
      Height = 13
      Caption = 'Nom du jeu de donn'#233'es :'
    end
    object lblLignesAAjouter: TLabel
      Left = 16
      Top = 143
      Width = 83
      Height = 13
      Caption = 'Ligne '#224' ajouter :'
    end
    object lblRowID: TLabel
      Left = 16
      Top = 162
      Width = 17
      Height = 13
      Caption = 'ID :'
    end
    object lblRowName: TLabel
      Left = 128
      Top = 162
      Width = 35
      Height = 13
      Caption = 'Name :'
    end
    object lblRowSurname: TLabel
      Left = 240
      Top = 162
      Width = 51
      Height = 13
      Caption = 'Surname :'
    end
    object edtWorkspaceID: TEdit
      Left = 16
      Top = 40
      Width = 320
      Height = 21
      TabOrder = 0
    end
    object edtDatasetID: TEdit
      Left = 16
      Top = 118
      Width = 320
      Height = 21
      TabOrder = 1
    end
    object edtDatasetAAjouter: TMemo
      Left = 16
      Top = 79
      Width = 320
      Height = 21
      TabOrder = 2
    end
    object edtRowID: TMemo
      Left = 16
      Top = 181
      Width = 96
      Height = 21
      TabOrder = 3
    end
    object edtRowName: TMemo
      Left = 128
      Top = 181
      Width = 96
      Height = 21
      TabOrder = 4
    end
    object edtRowSurname: TMemo
      Left = 240
      Top = 181
      Width = 96
      Height = 21
      TabOrder = 5
    end
  end
  object btnGetToken: TButton
    Left = 128
    Top = 322
    Width = 120
    Height = 30
    Caption = 'Obtenir le token'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold, fsUnderline]
    ParentFont = False
    TabOrder = 2
    OnClick = btnGetTokenClick
  end
  object grpSharePointActions: TGroupBox
    Left = 16
    Top = 353
    Width = 376
    Height = 125
    Caption = 'Actions SharePoint'
    TabOrder = 3
    object btnListSites: TButton
      Left = 8
      Top = 30
      Width = 168
      Height = 35
      Caption = 'Lister les sites'
      TabOrder = 0
      OnClick = btnListSitesClick
    end
    object btnListDirectories: TButton
      Left = 192
      Top = 30
      Width = 168
      Height = 35
      Caption = 'Lister les r'#233'pertoires d'#39'un site'
      TabOrder = 1
      OnClick = btnListDirectoriesClick
    end
    object btnListFiles: TButton
      Left = 8
      Top = 75
      Width = 168
      Height = 35
      Caption = 'Lister les fichiers'#13#10' d'#39'un r'#233'pertoire'
      TabOrder = 2
      WordWrap = True
      OnClick = btnListFilesClick
    end
    object btnUploadFile: TButton
      Left = 192
      Top = 75
      Width = 168
      Height = 35
      Caption = 'Envoyer un fichier'
      TabOrder = 3
      OnClick = btnUploadFileClick
    end
  end
  object grpPowerBIActions: TGroupBox
    Left = 416
    Top = 353
    Width = 376
    Height = 125
    Caption = 'Actions Power BI'
    TabOrder = 7
    object btnListWorkspaces: TButton
      Left = 8
      Top = 30
      Width = 168
      Height = 35
      Caption = 'Lister les espaces de travail'
      TabOrder = 0
      OnClick = btnListWorkspacesClick
    end
    object btnListDatasetsAndReports: TButton
      Left = 192
      Top = 30
      Width = 168
      Height = 35
      Caption = '  Lister les jeux de donn'#233'es et'#13#10'rapports d'#39'un espace de travail'
      TabOrder = 1
      WordWrap = True
      OnClick = btnListDatasetsAndReportsClick
    end
    object btnAddData: TButton
      Left = 8
      Top = 75
      Width = 168
      Height = 35
      Caption = 'Cr'#233'er un jeu de donn'#233'es'
      TabOrder = 2
      OnClick = btnAddDatasetClick
    end
    object btnAddRowsDataset: TButton
      Left = 192
      Top = 75
      Width = 168
      Height = 35
      Caption = 'Ajouter des donn'#233'es'
      TabOrder = 3
      OnClick = btnAddRowsDatasetClick
    end
  end
  object MemoOutput: TMemo
    Left = 16
    Top = 485
    Width = 776
    Height = 105
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 4
  end
end
