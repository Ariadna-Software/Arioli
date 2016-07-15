VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmppal 
   BackColor       =   &H00858585&
   Caption         =   "Ariges 4"
   ClientHeight    =   9315
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   13005
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageListB 
      Left            =   4920
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8688
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":909A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B8E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListPpal 
      Left            =   360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C2F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E418
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F4AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1053C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":13050
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":140E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":15174
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":16206
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":17298
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1832A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":193BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1A44E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1B4E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1C572
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1D604
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1E696
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1F728
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":210BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2791C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2BE1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2C830
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2D70A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   32
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artículos"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Movimientos Art."
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proveedores"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ofertas Clientes"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Clientes"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaranes Clientes"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Cliente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Proveedor"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaran Proveedor"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Factura Proveedor"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recepción Facturas Prov."
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimientos"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nº Serie"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gastos técnicos"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consulta precio artículo"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Venta TPV"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar empresa"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agenda"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Movimientos mercancia"
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   8730
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpal.frx":33F6C
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14870
            Text            =   "asdasd"
            TextSave        =   "asdasd"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:52"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   5640
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   49
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3752E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":39238
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3F4DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3FEF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":40902
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":430B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4398E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":44268
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":44B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4541C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":45E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":46288
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4639A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":464AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":465BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":468D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4C4FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4CF0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4D91E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4DA30
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4E442
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4EE54
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4F866
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4FE00
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5056C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":509BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":50E10
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":51262
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":516B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":51B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":51E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":51F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":52294
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":525AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":52E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":53762
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":53A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":53BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":53EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":54902
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":55314
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":55D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":56738
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5714A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":57B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5AF4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5B960
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5C372
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListTPV 
      Left            =   360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":62BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":64566
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":65EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6788A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6921C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6ABAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6C540
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6DED2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListMAIL 
      Left            =   360
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":74734
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":74B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":74FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7542A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7587C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":75CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":76120
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":76572
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":769C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7CC5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7D670
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8390A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8A16C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":909CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":97230
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9DA92
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A42F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AAB56
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AAFA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AB3FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AB84C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ABC9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AC0F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AC542
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B2164
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B2FB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B32D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B35EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B3904
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnConfiguracion 
      Caption         =   "C&onfiguración"
      Begin VB.Menu mnConfParamGenerales 
         Caption         =   "Datos &Empresa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnConfParamAplic 
         Caption         =   "Parámetros &Aplicación"
      End
      Begin VB.Menu mnConTMovimiento 
         Caption         =   "Tipos &Movimiento"
      End
      Begin VB.Menu mnConfParamRpt 
         Caption         =   "Tipos de &Documentos"
      End
      Begin VB.Menu mnAridoc1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnAridoc1 
         Caption         =   "Configuración aridoc"
         Index           =   1
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnConfManteUsuarios 
         Caption         =   "Mantenimiento &Usuarios"
         HelpContextID   =   2
      End
      Begin VB.Menu mnNuevaEmpresa 
         Caption         =   "Creacion &nueva empresa"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnUsuarios 
         Caption         =   "Nuevo U&suario"
         Visible         =   0   'False
      End
      Begin VB.Menu mnPedirPwd 
         Caption         =   "Password requerido"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCambioEmpresa 
         Caption         =   "Cambiar Em&presa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnBarra17 
         Caption         =   "-"
      End
      Begin VB.Menu mnSeleccionarImpresora 
         Caption         =   "Seleccionar &Impresora"
      End
      Begin VB.Menu mnBarra12 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnAlmacen 
      Caption         =   "&Almacen"
      Begin VB.Menu mnDatosGenAlmacen 
         Caption         =   "&Datos Generales"
         Begin VB.Menu mnAlmMarcas 
            Caption         =   "&Marcas"
         End
         Begin VB.Menu mnAlmAlPropios 
            Caption         =   "Almacenes &Propios"
         End
         Begin VB.Menu mnAlmTipoUnidad 
            Caption         =   "Tipos &Unidad"
         End
         Begin VB.Menu mnTiposArticulos 
            Caption         =   "&Tipos Articulos"
         End
         Begin VB.Menu mnAlmUbicacion 
            Caption         =   "U&bicaciones"
         End
         Begin VB.Menu mnAlmFamiliaArticulo 
            Caption         =   "&Familias Artículos"
         End
         Begin VB.Menu mnCtasContab 
            Caption         =   "Cuentas contabilizacion"
         End
         Begin VB.Menu mnAlmCategoria 
            Caption         =   "&Categorías"
         End
         Begin VB.Menu mnAlmArticulos 
            Caption         =   "&Artículos"
            Begin VB.Menu mnAlmArticulosMto 
               Caption         =   "Mantenimiento Articulos"
               Index           =   0
            End
            Begin VB.Menu mnAlmArticulosMto 
               Caption         =   "-"
               Index           =   1
            End
            Begin VB.Menu mnAlmArticulosMto 
               Caption         =   "Ficha técnica materia prima"
               Index           =   2
            End
         End
         Begin VB.Menu mnAlmNumLotes 
            Caption         =   "&Numeros de lote"
         End
      End
      Begin VB.Menu mnAlmMovimientosAlm 
         Caption         =   "&Movimientos Almacen"
         Begin VB.Menu mnAlmTraspaso 
            Caption         =   "&Traspaso Almacenes"
         End
         Begin VB.Menu mnAlmTraspasoHco 
            Caption         =   "&Histórico Traspaso Almacenes"
         End
         Begin VB.Menu mnAlmMovimientos 
            Caption         =   "&Movimientos Almacen"
         End
         Begin VB.Menu mnAlmMovimientosHco 
            Caption         =   "H&istórico Movimientos Almacen"
         End
         Begin VB.Menu mnbarracoupa 
            Caption         =   "-"
         End
         Begin VB.Menu mnCoupages 
            Caption         =   "Coupa&ges"
         End
         Begin VB.Menu mnMoixentMov 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnMoixentMov 
            Caption         =   "Movimientos mercancia"
            Index           =   1
         End
      End
      Begin VB.Menu mnAlmConsultas 
         Caption         =   "&Consultas"
         Begin VB.Menu mnAlmMovimArticulos 
            Caption         =   "Movimientos A&rticulos"
         End
         Begin VB.Menu mnAlmMovStock 
            Caption         =   "Movimientos desde inventario"
         End
         Begin VB.Menu mnAlmListMovim 
            Caption         =   "Listado &Movimientos"
         End
         Begin VB.Menu mnAlmListInactivos 
            Caption         =   "Listado Articulos &Inactivos"
         End
         Begin VB.Menu mnAlmListValoracion 
            Caption         =   "Listado Valoración &Stocks"
         End
         Begin VB.Menu mnAlmListMaxMin 
            Caption         =   "Inf. Stocks Máximos-Mínimos"
         End
         Begin VB.Menu mnAlmStockFecha 
            Caption         =   "Inf. Stocks a una &Fecha"
         End
      End
      Begin VB.Menu mnAlmInventario 
         Caption         =   "&Inventario"
         Begin VB.Menu mnAlmTomaInven 
            Caption         =   "&Toma de inventario"
         End
         Begin VB.Menu mnAlmEntradaInve 
            Caption         =   "&Entrada existencia real"
         End
         Begin VB.Menu mnAlmListadoInve 
            Caption         =   "&Listado diferencias"
         End
         Begin VB.Menu mnAlmActualizarInve 
            Caption         =   "Actualizar &direrencias"
         End
         Begin VB.Menu mnAlmValoracionInve 
            Caption         =   "&Valoración stocks inventariados"
         End
         Begin VB.Menu mnInventarioAceite 
            Caption         =   "Inventario aceite"
         End
         Begin VB.Menu mnBarra2 
            Caption         =   "-"
         End
         Begin VB.Menu mnAlmHcoInven 
            Caption         =   "&Histórico inventario"
         End
      End
   End
   Begin VB.Menu mnFacturacion 
      Caption         =   "&Facturación"
      Begin VB.Menu mnFacDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnFacActividades 
            Caption         =   "Activi&dades"
         End
         Begin VB.Menu mnFacZonas 
            Caption         =   "&Zonas"
         End
         Begin VB.Menu mnFacRutas 
            Caption         =   "&Rutas"
         End
         Begin VB.Menu mnFacFormasEnvio 
            Caption         =   "Formas de &Envio"
         End
         Begin VB.Menu mnFacFormasPago 
            Caption         =   "Formas de &Pago"
         End
         Begin VB.Menu mnFacBancosPropios 
            Caption         =   "&Bancos Propios"
         End
         Begin VB.Menu mnFacSituaciones 
            Caption         =   "&Situaciones Especiales"
         End
         Begin VB.Menu mnFacAgentesCom 
            Caption         =   "Agentes &Comerciales"
         End
         Begin VB.Menu mnFacClientesV1 
            Caption         =   "Clientes &Varios"
         End
         Begin VB.Menu mnFacClientes 
            Caption         =   "Cl&ientes"
         End
         Begin VB.Menu mnFacCartas 
            Caption         =   "Tipos de C&artas"
         End
         Begin VB.Menu mnFacIncidencias 
            Caption         =   "&Incidencias"
         End
      End
      Begin VB.Menu mnFacInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnFacInactivos 
            Caption         =   "Clientes Inacti&vos"
         End
         Begin VB.Menu mnFacInfClientes 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu mnFacAltas 
            Caption         =   "&Altas Clientes"
         End
         Begin VB.Menu mnFacEtiqClien 
            Caption         =   "&Etiquetas de clientes"
         End
         Begin VB.Menu mnFacCartaClien 
            Caption         =   "Car&tas a clientes"
         End
         Begin VB.Menu mnEtiquetasBultos 
            Caption         =   "&Etiquetas de bultos"
         End
      End
      Begin VB.Menu mnFacPreciosDtos 
         Caption         =   "&Precios y Descuentos"
         Begin VB.Menu mnFacTarifasVen 
            Caption         =   "&Tarifas Venta"
         End
         Begin VB.Menu mnFacTarifasPrecios 
            Caption         =   "&Lista Precios"
         End
         Begin VB.Menu mnFacPreEspecial 
            Caption         =   "Precios &Especiales"
         End
         Begin VB.Menu mnFacPromociones 
            Caption         =   "&Promociones"
         End
         Begin VB.Menu mnFacDescuentos 
            Caption         =   "&Descuentos Familia/Marca"
         End
         Begin VB.Menu mnFacBoniFact 
            Caption         =   "&Bonificaciones Factura"
         End
         Begin VB.Menu Barra12 
            Caption         =   "-"
         End
         Begin VB.Menu mnFactActPrecios 
            Caption         =   "&Actualizar precios"
         End
         Begin VB.Menu mnOliTarifasOfertas 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnOliTarifasOfertas 
            Caption         =   "Mantenimiento tarifas-ofertas"
            Index           =   1
         End
         Begin VB.Menu mnOliTarifasOfertas 
            Caption         =   "Mantenimiento tarifas"
            Index           =   2
         End
         Begin VB.Menu mnOliTarifasOfertas 
            Caption         =   "Generacion TOs"
            Index           =   3
         End
         Begin VB.Menu mnOliTarifasOfertas 
            Caption         =   "Copiar tarifa"
            Index           =   4
         End
         Begin VB.Menu mnOliTarifasOfertas 
            Caption         =   "Historico TOs-Tarifas"
            Index           =   5
         End
         Begin VB.Menu mnbarra39 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacInfMargenes 
            Caption         =   "&Control margenes tarifas"
         End
         Begin VB.Menu mnPreciosTarifasCorreccion 
            Caption         =   "Corrección errores y actualización tarifas"
         End
         Begin VB.Menu mnAVABprecio 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnAVABprecio 
            Caption         =   "Actualizar precios desde proveedor"
            Index           =   1
         End
         Begin VB.Menu mnRecalculoPRecioSt 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnRecalculoPRecioSt 
            Caption         =   "Recálculo precio standard"
            Index           =   1
         End
      End
      Begin VB.Menu mnFacOfert 
         Caption         =   "&Ofertas"
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Mantenimiento Ofertas"
            Index           =   0
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Grupo de Plantillas"
            Index           =   1
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "Entrada de  &Plantillas"
            Index           =   2
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "Ofertas E&fectuadas"
            Index           =   3
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Histórico  Ofertas"
            Index           =   5
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Traspaso a Histórico"
            Index           =   6
         End
      End
      Begin VB.Menu mnFacPed 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Mantenimiento Pedidos"
            Index           =   0
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Histórico Pedidos Anulados"
            Index           =   1
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Cartas Confirmacion de Pedidos"
            Index           =   3
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe &Pedidos por Articulo"
            Index           =   4
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe P&edidos por Cliente"
            Index           =   5
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe &Disponibilidad Stocks"
            Index           =   6
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Consulta precios / cliente"
            Index           =   8
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Estadísticas consultas pedidos"
            Index           =   9
         End
      End
      Begin VB.Menu mnFacAlbaran 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnFacEntAlbaran 
            Caption         =   "&Mantenimiento Albaranes"
         End
         Begin VB.Menu mnAlbaranesB 
            Caption         =   "Albaranes presupuestos *"
         End
         Begin VB.Menu mnFacAlbxArtic 
            Caption         =   "Informe &Albaranes por Articulo"
         End
         Begin VB.Menu mnFacIncumPlazos 
            Caption         =   "Inf. Incumplimiento Plazos &Ent."
         End
         Begin VB.Menu mnFacHcoAlbaranes 
            Caption         =   "&Histórico Albaranes Anulados"
         End
         Begin VB.Menu mnBarra5 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacPreFacturar 
            Caption         =   "&Previsión Facturación"
         End
         Begin VB.Menu mnFacFacturarAlb 
            Caption         =   "&Facturación de Albaranes"
         End
         Begin VB.Menu mnFacAlbMostrador 
            Caption         =   "Facturas de Mo&strador"
         End
         Begin VB.Menu mnFacturarPresupuestos 
            Caption         =   "Facturar presupuestos *"
         End
         Begin VB.Menu mnFacAlbRectifica 
            Caption         =   "Facturas &Rectificativas"
         End
         Begin VB.Menu mnFacHcoFacturas 
            Caption         =   "His&tórico Albaran/Factura"
         End
         Begin VB.Menu mnFacReImpFactu 
            Caption         =   "Re&imprimir Facturas"
         End
         Begin VB.Menu mnEnvioFactuasMail 
            Caption         =   "Enviar facturas por e&mail"
         End
         Begin VB.Menu mnExportarFacturasPDF 
            Caption         =   "Exportar facturas PDF"
         End
         Begin VB.Menu mnServicios 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Albaranes de servicio"
            Index           =   1
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Facturación de servicios"
            Index           =   2
         End
         Begin VB.Menu mnServicios 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Albaranes internos"
            Index           =   4
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Facturación albaranes internos"
            Index           =   5
         End
         Begin VB.Menu mnTicket 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnTicket 
            Caption         =   "Contabilizar facturas tickets agrupados"
            Index           =   1
         End
         Begin VB.Menu mnTicket 
            Caption         =   "Listado tickets facturados"
            Index           =   2
         End
         Begin VB.Menu mnTraspasoFraAVAB 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnTraspasoFraAVAB 
            Caption         =   "Traspaso facturas AVAB"
            Index           =   1
         End
         Begin VB.Menu mnBarra9 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
      End
      Begin VB.Menu mnRutas_ 
         Caption         =   "Ordenes de carga"
         Begin VB.Menu mnRutas 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu mnRutas 
            Caption         =   "Vehiculos"
            Index           =   1
         End
      End
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnFacEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnFacEstVentaAceite 
            Caption         =   "Ventas &aceite"
         End
         Begin VB.Menu mnFacEstVentaCliente 
            Caption         =   "&Ventas por cliente"
         End
         Begin VB.Menu mnFacEstVentaTraba 
            Caption         =   "Ventas por &trabajador"
         End
         Begin VB.Menu mnFacEstVentaMes 
            Caption         =   "Ventas por &meses"
         End
         Begin VB.Menu mnFacEstVentaFam 
            Caption         =   "Ventas por &familia  /  Artículo"
         End
         Begin VB.Menu mnVentasPorProveedor 
            Caption         =   "Ventas por proveedor"
         End
         Begin VB.Menu mnVentasAgente 
            Caption         =   "Ventas por agente"
         End
         Begin VB.Menu mnEcoenves 
            Caption         =   "ECOEMBES"
         End
         Begin VB.Menu mnFacEstDetalleFac 
            Caption         =   "&Detalle facturación"
         End
         Begin VB.Menu mnVtasAgrupadox 
            Caption         =   "Ventas agrupado por..."
         End
         Begin VB.Menu mnFacEstMargenVtas 
            Caption         =   "Mar&gen ventas por artículo "
         End
      End
   End
   Begin VB.Menu mnCompras 
      Caption         =   "&Compras"
      Begin VB.Menu mnComDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnComProveedores 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnComProveVarios 
            Caption         =   "Proveedores &Varios"
         End
         Begin VB.Menu mnComDirecciones 
            Caption         =   "&Direcciones"
         End
      End
      Begin VB.Menu mnComInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnComInfProve 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnComEtiqProve 
            Caption         =   "&Etiquetas de proveedores"
         End
         Begin VB.Menu mnComCartaProve 
            Caption         =   "&Cartas a Proveedores"
         End
      End
      Begin VB.Menu mnComPreciosDtos 
         Caption         =   "Precios y &Descuentos"
         Begin VB.Menu mnComPreProve 
            Caption         =   "P&recios Proveedor"
            Index           =   0
         End
         Begin VB.Menu mnComPreProve 
            Caption         =   "Descuentos Pro&veedor"
            Index           =   1
         End
         Begin VB.Menu mnComPreProve 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnComPreProve 
            Caption         =   "Actualizar precios"
            Index           =   3
         End
      End
      Begin VB.Menu mnComPedidos 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnComPedMant 
            Caption         =   "Mant. &Pedidos Proveedor"
         End
         Begin VB.Menu mnComHcoPedidos 
            Caption         =   "&Histórico Pedidos Anulados"
         End
         Begin VB.Menu mnComPteRecibir 
            Caption         =   "List. &Material pendiente de recibir"
         End
      End
      Begin VB.Menu mnComAlbaranes 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnComAlbMan 
            Caption         =   "&Mant. Albaranes Proveedor"
         End
         Begin VB.Menu mnComHcoAlbaranes 
            Caption         =   "&Histórico Albaranes Anulados"
         End
         Begin VB.Menu mnComPteFacturar 
            Caption         =   "List. &Pendiente de facturar"
         End
         Begin VB.Menu mnBarra7 
            Caption         =   "-"
         End
         Begin VB.Menu mnComFacturar 
            Caption         =   "&Recepción Facturas"
         End
         Begin VB.Menu mnComHcoFacturas 
            Caption         =   "&Histórico Albaran/Factura"
         End
         Begin VB.Menu mnBarra15 
            Caption         =   "-"
         End
         Begin VB.Menu mnComContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
      End
      Begin VB.Menu mnProcesoLiquidacionProveedores 
         Caption         =   "Liquidación proveedores"
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Cambiar precios"
            Index           =   0
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Liquidacion proveedores"
            Index           =   1
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Impresion facturas"
            Index           =   2
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Asociar albaranes compras / ventas"
            Index           =   3
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Listado asociaciones albaranes"
            Index           =   4
         End
      End
      Begin VB.Menu Barra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnComEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnComEstComprasxProve 
            Caption         =   "Compras por &Proveedor"
         End
         Begin VB.Menu mnComEstComprasxFam 
            Caption         =   "Compras por &Familia/Artíc."
         End
         Begin VB.Menu mnComEstAlbarxProve 
            Caption         =   "&Albaranes por Proveedor"
         End
      End
   End
   Begin VB.Menu mnAdministracion 
      Caption         =   "A&dministración"
      Begin VB.Menu mnAdmDatosGen 
         Caption         =   "&Datos Generales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnAdmTrabajadores 
         Caption         =   "&Trabajadores"
      End
      Begin VB.Menu mnAdmGastosTec 
         Caption         =   "&Gastos Técnicos"
      End
      Begin VB.Menu mnAdmNominas 
         Caption         =   "&Nominas y Gastos"
      End
   End
   Begin VB.Menu mnMantenimientos 
      Caption         =   "&Mantenimientos"
      Visible         =   0   'False
      Begin VB.Menu mnManTiposContrato 
         Caption         =   "&Tipos de Contrato"
      End
      Begin VB.Menu mnManEntrada 
         Caption         =   "&Entrada Mantenimientos"
      End
      Begin VB.Menu mnBarra8 
         Caption         =   "-"
      End
      Begin VB.Menu mnManListado 
         Caption         =   "&Listado Mantenimientos"
      End
      Begin VB.Menu mnManRevisiones 
         Caption         =   "Listado &Revisiones Mant."
      End
      Begin VB.Menu mnManFichas 
         Caption         =   "&Fichas Mantenimientos"
      End
      Begin VB.Menu mnManAltas 
         Caption         =   "List. &Altas Mantenimientos"
      End
      Begin VB.Menu mnInfTeoMant 
         Caption         =   "Informe teórico mantenimientos"
      End
      Begin VB.Menu mnEtiqMante 
         Caption         =   "Etiquetas de mantenimientos"
      End
      Begin VB.Menu mnBarra30 
         Caption         =   "-"
      End
      Begin VB.Menu mnCartaRenovaMante 
         Caption         =   "Carta renovación"
      End
      Begin VB.Menu mnTraspasoMante 
         Caption         =   "Traspaso siguiente a actual"
      End
      Begin VB.Menu mnBarra32 
         Caption         =   "-"
      End
      Begin VB.Menu mnHcoMaten 
         Caption         =   "Histórico mantenimientos anulados"
      End
      Begin VB.Menu mnInfManteAnulados 
         Caption         =   "Informe mantenimientos anulados"
      End
      Begin VB.Menu mnBarra13 
         Caption         =   "-"
      End
      Begin VB.Menu mnManPrevFac 
         Caption         =   "&Previsión Facturación"
      End
      Begin VB.Menu mnManFactAlb 
         Caption         =   "Fac&turación  Mantenimientos"
      End
   End
   Begin VB.Menu mnReparaciones 
      Caption         =   "&Reparaciones"
      Visible         =   0   'False
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "&Mant.  Reparaciones"
         Index           =   0
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "C&ontrol Reparaciones"
         Index           =   1
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Mant. &Nº Serie"
         Index           =   2
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Motivos &baja equipos"
         Index           =   3
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Motivos &Pend. Rep."
         Index           =   4
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "&Histórico de Reparaciones"
         Index           =   5
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Servicios asistencia técnica"
         Index           =   6
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Tipos averia"
         Index           =   7
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Trabajos realizados"
         Index           =   8
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Listado Rep. del &Dia"
         Index           =   10
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Listado Rep. por &Cliente"
         Index           =   11
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "F&recuencia de reparaciones"
         Index           =   12
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Estadística reparaciones técnico"
         Index           =   13
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Listado reparaciones efectuadas"
         Index           =   14
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Mant. &Albaranes Rep."
         Index           =   16
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "Pre&visión Facturación"
         Index           =   17
      End
      Begin VB.Menu mnRepEntReparacion2 
         Caption         =   "&Facturación Reparaciones"
         Index           =   18
      End
   End
   Begin VB.Menu mnproduccion 
      Caption         =   "Producción"
      Begin VB.Menu mnproduccion1 
         Caption         =   "Órdenes producción"
         Index           =   0
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Descripción costes "
         Index           =   1
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Lineas produccion"
         Index           =   3
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "NO VISIBLE. DISPONIBLE"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Trazabilidad / Producción"
         Index           =   5
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Depositos"
         Index           =   6
         Begin VB.Menu mnproduccion1_1 
            Caption         =   "Mantenimiento "
            Index           =   0
         End
         Begin VB.Menu mnproduccion1_1 
            Caption         =   "Trasiegos"
            Index           =   1
         End
         Begin VB.Menu mnproduccion1_1 
            Caption         =   "Filtrado"
            Index           =   2
         End
         Begin VB.Menu mnproduccion1_1 
            Caption         =   "Vaciado depósito"
            Index           =   3
         End
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Mantenimiento palets"
         Index           =   8
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Mantenimiento etiquetas"
         Index           =   9
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Pistola"
         Index           =   11
      End
      Begin VB.Menu mnNuevosPuntosMenuTraza 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnNuevosPuntosMenuTraza 
         Caption         =   "Control calidad"
         Index           =   1
      End
      Begin VB.Menu mnNuevosPuntosMenuTraza 
         Caption         =   "Partidas"
         Index           =   2
         Begin VB.Menu mnPartida2 
            Caption         =   "Mantenimientos "
            Index           =   0
         End
         Begin VB.Menu mnPartida2 
            Caption         =   "Movimientos "
            Index           =   1
         End
      End
      Begin VB.Menu mnNuevosPuntosMenuTraza 
         Caption         =   "Trazabilidad"
         Index           =   3
         Begin VB.Menu mnTrazaNueva 
            Caption         =   "Lote venta"
            Index           =   0
         End
         Begin VB.Menu mnTrazaNueva 
            Caption         =   "Desde compra"
            Index           =   1
         End
         Begin VB.Menu mnTrazaNueva 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnTrazaNueva 
            Caption         =   "Desde venta"
            Index           =   3
         End
         Begin VB.Menu mnTrazaNueva 
            Caption         =   "Origen facturas-albaranes"
            Index           =   4
         End
         Begin VB.Menu mnTrazaNueva 
            Caption         =   "Destino lotes"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnAlmazara 
      Caption         =   "Almazara"
      Begin VB.Menu mnAlmazara1 
         Caption         =   "Entrada oliva"
         Index           =   0
      End
      Begin VB.Menu mnAlmazara1 
         Caption         =   "Proceso en almazara"
         Index           =   1
      End
   End
   Begin VB.Menu mnTPV 
      Caption         =   "&Punto de Venta"
      Visible         =   0   'False
      Begin VB.Menu mnTPV2 
         Caption         =   "Pantalla de &venta"
         Index           =   0
      End
      Begin VB.Menu mnTPV2 
         Caption         =   "&Cierre de caja"
         Index           =   1
      End
      Begin VB.Menu mnTPV2 
         Caption         =   "Etiquetas estantería"
         Index           =   2
      End
      Begin VB.Menu mnTPV2 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnTPV2 
         Caption         =   "&Parámetros generales TPV"
         Index           =   4
      End
      Begin VB.Menu mnTPV2 
         Caption         =   "Parámetros &terminales TPV"
         Index           =   5
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnAgenda 
         Caption         =   "&Agenda"
      End
      Begin VB.Menu mnVerAvisos 
         Caption         =   "A&visos"
      End
      Begin VB.Menu mnBackUp 
         Caption         =   "&Copia Seguridad local"
      End
      Begin VB.Menu mnRecupFac 
         Caption         =   "&Recuperar facturas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminarFacturas 
         Caption         =   "&Borre Facturas y Movimientos"
      End
      Begin VB.Menu mnRevisarMultibase 
         Caption         =   "Revisar caracteres especiales"
      End
      Begin VB.Menu mnManteneLOG 
         Caption         =   "Acciones realizadas"
      End
      Begin VB.Menu mnAridocFacturas 
         Caption         =   "Traspaso Aridoc"
      End
      Begin VB.Menu mnUtiDeclaracionLOM 
         Caption         =   "Declaración &LOM"
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminarArticulos 
         Caption         =   "Eliminar articulos"
      End
      Begin VB.Menu mnCambiarArticuloFacturado 
         Caption         =   "Cambiar articulo fac."
      End
      Begin VB.Menu mnEtiquetasArticulos 
         Caption         =   "Etiquetas artículos"
      End
      Begin VB.Menu mnRegistros2 
         Caption         =   "-"
      End
      Begin VB.Menu mnRegPpañ 
         Caption         =   "Registros"
         Begin VB.Menu mnRegistros 
            Caption         =   "Limpieza y desinfeccion"
            Index           =   1
         End
         Begin VB.Menu mnRegistros 
            Caption         =   "Mantenimiento preventivo"
            Index           =   2
         End
         Begin VB.Menu mnRegistros 
            Caption         =   "Listado revisiones"
            Index           =   3
         End
         Begin VB.Menu mnRegistros 
            Caption         =   "Cloro"
            Index           =   4
         End
         Begin VB.Menu mnRegistros 
            Caption         =   "Mantenimiento lista revision"
            Index           =   5
         End
         Begin VB.Menu mnRegistros 
            Caption         =   "Acciones correctoras"
            Index           =   6
         End
         Begin VB.Menu mnRegistros 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnRegistros 
            Caption         =   "Documentos pendientes revisar"
            Index           =   8
         End
      End
      Begin VB.Menu mnBarra19 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiBuscar 
         Caption         =   "&Buscar..."
         Begin VB.Menu mnUtiBuscarErrFac 
            Caption         =   "&Errores en Nº Factura clientes"
         End
         Begin VB.Menu mnUtiBuscarPteCon 
            Caption         =   "Facturas pendientes de &contabilizar"
            Begin VB.Menu mnUtiBuscarErrConCli 
               Caption         =   "&Clientes"
            End
            Begin VB.Menu mnUtiBuscarErrConPro 
               Caption         =   "&Proveedores"
            End
         End
      End
      Begin VB.Menu mnBarra20 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiUsuActivos 
         Caption         =   "&Usuarios activos"
      End
   End
   Begin VB.Menu mnSoporte2 
      Caption         =   "&Soporte"
      Begin VB.Menu mnSoporte 
         Caption         =   "Ayuda"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Enviar Mail"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Web Ariadna Software"
         Index           =   4
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Comprobar version operativa"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Cambiar password"
         Index           =   7
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Acerca de ..."
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrimeraVez As Boolean

Dim TieneEditorDeMenus As Boolean


Private Sub MDIForm_Activate()
Dim b As Boolean
Dim I As Integer
Dim Permis As String

'Dim AvisosPendientes As Boolean
'Formulario Principal
   ' AvisosPendientes = False
    I = 0
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
       ' AvisosPendientes = TieneAvisosPendientes()
       ComprobarDocumentosPendientes True
       I = 1
       
    End If
    If Not vParam Is Nothing Then
        If vParam.Modificado Then
          'Poner datos visible del form
           PonerDatosVisiblesForm
           vParam.Modificado = False
        End If
    End If
    
    If I = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '-- Control de si se utilizan servicios o no ( si es que no no se muestra el menú)
    '   el situarlo aqui hace que no haya que salir y entrar en el programa si se
    mnServicios(0).visible = vParamAplic.Servicios
    mnServicios(1).visible = vParamAplic.Servicios
    mnServicios(2).visible = vParamAplic.Servicios
    
    
    'Albaranes INTERNOS
    b = DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "ALI", "T") <> ""
    PuntoDeMenuVisible mnServicios(3), b
    PuntoDeMenuVisible mnServicios(4), b
    PuntoDeMenuVisible mnServicios(5), b
    
    
    'vParamAplic.Reparaciones
    
    'MAntenimientos y reparaciones
    mnMantenimientos.visible = vParamAplic.Mantenimientos
    mnReparaciones.visible = vParamAplic.Reparaciones
    
    '-- Eliminamos frecuencias de momento
    'mnFrecuencias.visible = vParamAplic.Frecuencias
    'Me.mnbarra33.visible = mnFrecuencias.visible
    
    '--------------------

    '-- Contabilizacion tickets agrupados
    mnTicket(0).visible = vParamAplic.ContabilizarTicketAgrupados
    mnTicket(1).visible = vParamAplic.ContabilizarTicketAgrupados
    mnTicket(2).visible = vParamAplic.ContabilizarTicketAgrupados
    
        
    'Los albaranes y facturas en "B"
    'seran visibles si esta creado el tipo movimiento y tene contabilidad B
    'b = DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "ALZ", "T") <> ""
    'b = b And vParamAplic.ContabilidadB > 0
    b = vUsu.TrabajadorB
    PuntoDeMenuVisible Me.mnAlbaranesB, b
    PuntoDeMenuVisible mnFacturarPresupuestos, b
    
    'Declara LOM visible : FALSE
    PuntoDeMenuVisible mnUtiDeclaracionLOM, False
    'De momento:
    PuntoDeMenuVisible Me.mnAridoc1(0), True
    PuntoDeMenuVisible Me.mnAridoc1(1), True
    
    'PuntoDeMenuVisible mnTelefonia, False
    
    'Solo administradores
    PuntoDeMenuVisible mnNuevosPuntosMenuTraza(1), (vUsu.Nivel < 2)
    PuntoDeMenuVisible mnNuevosPuntosMenuTraza(2), (vUsu.Nivel < 2)
       
    'Produccion
    PuntoDeMenuVisible Me.mnproduccion, vParamAplic.Produccion
       
    'Registros de limpieza, desinfeccion.... Si produccion y tiene AVAB.
    b = vParamAplic.Produccion And EmprAVAB > 0
    PuntoDeMenuVisible mnRegPpañ, b
    PuntoDeMenuVisible mnRegistros2, b
    'PuntoDeMenuVisible mnRegistros2, vParamAplic.Produccion
    
    'Actualizar precios en AVAB desde morales
    'B = vEmpresa.codempre = EmpresaAVAB
    b = vParamAplic.EsAVAB
    PuntoDeMenuVisible mnAVABprecio(0), b
    PuntoDeMenuVisible mnAVABprecio(1), b
    PuntoDeMenuVisible mnTraspasoFraAVAB(0), b
    PuntoDeMenuVisible mnTraspasoFraAVAB(1), b
    
        
    mnRutas(0).visible = vParamAplic.QUE_EMPRESA <> 4
        
    mnproduccion1(4).visible = False   'Disponible para poner
    mnAlmazara.visible = False
    
    'Si no tiene NUEVA PRODUCCION no dejamos ver los puntos de las lineas...
    'El 4 de momento NO ES VISIBLE, esta disponible. Empiezo en el 5
    If vParamAplic.QUE_EMPRESA = 0 Then
    
        '0.-Produccion antigua
        '1.- Tasas
        '3.- Nueva produccion. Las lineas
        '5.- Ver trazabilidades hco
        '6.- caja, esta siempre visible=false
        '8.- palets
        '9.- etiquetas
        '11.- pistola
       
        
        b = vParamAplic.ProduccionNueva And vUsu.Nivel = 0
        PuntoDeMenuVisible mnproduccion1(9), b   'etiquetas solo ramon y adninistrador
     
    Else
        'Caulqueir otra NO
        For I = 2 To Me.mnproduccion1.Count - 1
            PuntoDeMenuVisible mnproduccion1(I), False
        Next
    
        If vParamAplic.QUE_EMPRESA = 4 Then    'QUATRETONDA ALMAZARA
    
            '0.-Produccion antigua
            '1.- Tasas
            '3.- Nueva produccion. Las lineas
            '5.- Ver trazabilidades hco
            '6.- caja, esta siempre visible=false
            '8.- palets
            '9.- etiquetas
            '11.- pistola
           
            

            mnproduccion1(6).visible = True
            mnproduccion1(1).visible = False
            mnNuevosPuntosMenuTraza(1).visible = False
    

            mnAlmazara.visible = True

        End If
        
    End If
    
    
    
    '-- Descriptores especiales (Vrs 4.0.9)
    If vParamAplic.Descriptores Then
        mnAlmTipoUnidad.Caption = "Formatos"
        mnTiposArticulos.Caption = "Modelos"
        mnAlmFamiliaArticulo.Caption = "Categorias Art."
        mnAlmCategoria.visible = False
    End If
    
    PuntoDeMenuVisible mnCtasContab, vParamAplic.ContabilizacionMoixent
    
    PuntoDeMenuVisible Me.mnMoixentMov(0), (vParamAplic.QUE_EMPRESA = 2 Or vParamAplic.QUE_EMPRESA = 3)
    PuntoDeMenuVisible Me.mnMoixentMov(1), (vParamAplic.QUE_EMPRESA = 2 Or vParamAplic.QUE_EMPRESA = 3)
    
    'De momento esta  oculto
    'PuntoDeMenuVisible Me.mnMoixentMov(0), False
    'PuntoDeMenuVisible Me.mnMoixentMov(1), False
    Me.Toolbar1.Buttons(30).visible = Me.mnMoixentMov(0).visible
    
    
    'Estos LOTES son de ariges, no de arioli.
    'Lo oculto para evitar problemas
    PuntoDeMenuVisible mnAlmNumLotes, False
    
    
    PuntoDeMenuVisible mnTPV, False
    
    
    '--
    Screen.MousePointer = vbDefault
End Sub




Private Sub PuntoDeMenuVisible(ByRef MnPuntoDMenu As Menu, b As Boolean)
    If MnPuntoDMenu.visible Then MnPuntoDMenu.visible = b
    
End Sub


Private Sub MDIForm_Load()
'Formulario Principal

    CargaImagen

    PrimeraVez = True
    'Botones
    With Me.Toolbar1
        .ImageList = Me.ImgListPpal
        .Buttons(1).Image = 1   'Articulos
        .Buttons(2).Image = 2   'Movimientos Articulos
        
        .Buttons(5).Image = 3   'Clientes
        .Buttons(6).Image = 4   'Proveedores

        .Buttons(9).Image = 5   'Ofertas Clientes
        .Buttons(10).Image = 6   'Pedidos Clientes
        .Buttons(11).Image = 7   'Albaranes Clientes
        .Buttons(12).Image = 8   'Hist. Albaranes Clientes (Facturas)

        .Buttons(15).Image = 9   'Pedidos Proveedor
        .Buttons(16).Image = 10   'Albaranes Proveedor
        .Buttons(17).Image = 11   'Facturas Proveedor
        .Buttons(18).Image = 12   'Recepcion Facturas Proveedor
        
        .Buttons(21).Image = 15   'Mantenimientos
        .Buttons(22).Image = 16   'Nº Serie
        
        
        .Buttons(24).Image = 13 'Gastos tecnicos
        .Buttons(25).Image = 22 'Consulta precio articulo
        .Buttons(26).Image = 19 'Pantalla venta del TPV
        .Buttons(27).Image = 21 'Agenda
        .Buttons(28).Image = 20 'Agenda
        
        .Buttons(30).Image = 24 'Moviemiento para Moixent. Generara un albaran o un regularizacion
        
        .Buttons(32).Image = 14 'Salir
    End With
    LeerEditorMenus
    PonerDatosFormulario False
    
       
    'Fijar primer dia la semana en vbMyMonday
    'Para el calendario.
    FijarPrimerDiaSemana
    
    
    

       
    
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\arifon4.dat")
    If Err.Number <> 0 Then
        Me.Picture = LoadPicture()
        Err.Clear
    End If
End Sub




Private Sub PonerDatosFormulario(DesdeCambiarEmpresa As Boolean)
Dim Config As Boolean


    If Not DesdeCambiarEmpresa Then
        Config = (vEmpresa Is Nothing) Or (vParam Is Nothing) Or (vParamAplic Is Nothing)
    
        If Config Then HabilitarSoloPrametros_o_Empresas False
    End If
    
    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If DesdeCambiarEmpresa Then
        ReestablecerMenus
        HabilitarSoloPrametros_o_Empresas True
    End If
    
    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    
    'Comprobar que los iconos de la barra su correspondiente
    'entrada de menu esta habilitada sino desabilitar
    PoneBarraMenus
    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'Formulario Principal
Dim cad As String

    'Alguna cosilla antes de cerrar. Eliminar bloqueos
    cad = "Delete from zbloqueos where codusu = " & vUsu.Codigo
    conn.Execute cad

    'Elimnar bloquo BD
    Set vUsu = Nothing
    Set vConfig = Nothing
    Set vEmpresa = Nothing
    
    Set vParam = Nothing
    Set vParamAplic = Nothing
    
    
    TerminaBloquear
    
    'cerrar las conexiones
    conn.Close
    CerrarConexionConta

End Sub




Private Sub mnAdmGastosTec_Click()
'Gastos Técnicos
    frmAdmGasTec.Show vbModal
End Sub

Private Sub mnAdmNominas_Click()
'Nominas y Gastos
    frmAdmNominas.Show vbModal
End Sub

Private Sub mnAdmTrabajadores_Click()
    frmAdmTrabajadores.Show vbModal
End Sub

Private Sub mnAgenda_Click()

    MsgBox "Se ha producido un error abriendo la agenda", vbExclamation
    'frmMainCalendar.Show
End Sub

Private Sub mnAlbaranesB_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALZ"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnAlmActualizarInve_Click()
    AbrirListado (14)
End Sub

Private Sub mnAlmAlPropios_Click()
    frmAlmAlPropios.Show vbModal
End Sub

Private Sub mnAlmArticulosMto_Click(Index As Integer)
    Select Case Index
    Case 0
      frmAlmArticulos.Show vbModal
    
    Case 2
        'Ficha tencinca
        frmFichaTecnicaMP.Show vbModal
    End Select
End Sub

Private Sub mnAlmazara1_Click(Index As Integer)
    Select Case Index
    Case 0
        frmVallEntradaOliva.Show vbModal
       
    Case 1
    
    End Select
End Sub

Private Sub mnAlmCategoria_Click()
    'categorias de articulos
    frmAlmCategorias.Show vbModal
End Sub

Private Sub mnAlmEntradaInve_Click()
    frmAlmInventario.Show vbModal
End Sub

Private Sub mnAlmFamiliaArticulo_Click()
    frmAlmFamiliaArticulo.Show vbModal
End Sub


Private Sub mnAlmHcoInven_Click()
    frmAlmHcoInven.Show vbModal
End Sub

Private Sub mnAlmListadoInve_Click()
    AbrirListado (13)
End Sub

Private Sub mnAlmListInactivos_Click()
    AbrirListado (15)
End Sub

Private Sub mnAlmListMaxMin_Click()
'Informe de Stocks Maximos y Minimos
    AbrirListado (18)
End Sub

Private Sub mnAlmListMovim_Click()
    AbrirListado (9)
End Sub

Private Sub mnAlmListValoracion_Click()
    AbrirListado (17)
End Sub

Private Sub mnAlmMarcas_Click()
    frmAlmMarcas.Show vbModal
End Sub

Private Sub mnAlmMovimArticulos_Click()
    frmAlmMovimArticulos.Show vbModal
End Sub

Private Sub mnAlmMovimientos_Click()
    frmAlmMovimientos2.EsHistorico = False
    frmAlmMovimientos2.TrasiegoMoixent = False
    frmAlmMovimientos2.hcoCodMovim = -1 'No carga el form al abrir
    frmAlmMovimientos2.Show vbModal
End Sub

Private Sub mnAlmMovimientosHco_Click()
    frmAlmMovimientos2.EsHistorico = True
    frmAlmMovimientos2.TrasiegoMoixent = False
    frmAlmMovimientos2.hcoCodMovim = -1
    frmAlmMovimientos2.Show vbModal
End Sub

Private Sub mnAlmMovStock_Click()
    frmAlmMovArtSaldo.Show vbModal
End Sub

Private Sub mnAlmNumLotes_Click()
'numero de lote de los artículos
    frmAlmNumLote.Show vbModal
End Sub

Private Sub mnAlmStockFecha_Click()
'Informe de Stocks a una Fecha
    AbrirListado (19)
End Sub

Private Sub mnAlmTipoUnidad_Click()
    frmAlmTipoUnidad.Show vbModal
End Sub

Private Sub mnAlmTomaInven_Click()
    AbrirListado (12)
End Sub

Private Sub mnAlmTraspaso_Click()
    frmAlmTraspaso.EsHistorico = False
    frmAlmTraspaso.hcoCodMovim = -1
    frmAlmTraspaso.Show vbModal
End Sub

Private Sub mnAlmTraspasoHco_Click()
    frmAlmTraspaso.EsHistorico = True
    frmAlmTraspaso.hcoCodMovim = -1
    frmAlmTraspaso.Show vbModal
End Sub

Private Sub mnAlmUbicacion_Click()
    frmAlmUbicaciones.Show vbModal
End Sub

Private Sub mnAlmValoracionInve_Click()
    AbrirListado (16)
End Sub



Private Sub mnAridoc1_Click(Index As Integer)


    'Configuracion aridoc
    If Index = 1 Then HacerMenuARidoc 0
    
End Sub

Private Sub mnAridocFacturas_Click()
    frmAridocSeleccion.vOpcion = 1
    frmAridocSeleccion.Show vbModal
End Sub

Private Sub mnAVABprecio_Click(Index As Integer)
    'frmListado.OpcionListado = 510
    'frmListado.Show vbModal
    AbrirListado 510
End Sub

Private Sub mnBackUp_Click()
'Copia de seguridad de toda la base de datos
    frmBackUP.Show vbModal
End Sub





Private Sub mnCambiarArticuloFacturado_Click()
    frmFacCambiaArticulo.Show vbModal
    
End Sub

Private Sub mnCambioEmpresa_Click()
    Dim AntUSU As Usuario

    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If

    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.Codigo


    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"
    
'    Set AntUSU = vUsu
'    Set vUsu = Nothing
    frmLogin.Show vbModal
'    If vUsu Is Nothing Then
'        Set vUsu = AntUSU
'        Set AntUSU = Nothing
'        Exit Sub
'    End If

    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    conn.Close
    ConnConta.Close


    'Abre la conexión a BDatos:Ariges
    If AbrirConexion() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    Else
        Set vParam = Nothing
        Set vParamAplic = Nothing
        'Carga Parametros Generales y Contables de la empresa
        LeerParametros
    End If


    'Abrir conexión a la BDatos de Contabilidad para acceder a
    'Tablas: Cuentas, Tipos IVA
    If AbrirConexionConta(False) = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
        End
    End If

    
        'Trabajador que mete en ALMACEN B
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "straba", "presupuesto", "login", vUsu.Login, "T")
    vUsu.TrabajadorB = CadenaDesdeOtroForm = "1"  'Trabajador de almacen en B
    vUsu.FijarCodigoTrabajador
    
    
    Set vEmpresa = Nothing
    'LeerEmpresaParametros
    
     'Carga los Datos Básicos de la empresa
    LeerDatosEmpresa
    
    
    'Carga los Niveles de cuentas de Contabilidad de la empresa
    LeerNivelesEmpresa
    
    
    
    
'    PonerDatosFormulario
    PonerDatosVisiblesForm

    'Ponemos primera vez a false
    PonerDatosFormulario True
    PrimeraVez = True
    MDIForm_Activate

    

    Screen.MousePointer = vbDefault
End Sub


Private Sub mnCartaRenovaMante_Click()
    AbrirListado 78
End Sub

'Private Sub mnCheckVersion_Click()
''    Screen.MousePointer = vbHourglass
''    LanzaHome "webversion"
''    espera 2
''    Screen.MousePointer = vbDefault
'End Sub


Private Sub mnComAlbMan_Click()
'Mantenimiento de Albaranes a Proveedor
    frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmComEntAlbaranes.EsHistorico = False
    frmComEntAlbaranes.Show vbModal
End Sub

Private Sub mnComCartaProve_Click()
'Cartas a proveedores
     AbrirListadoOfer (306) '306: Informe Cartas a Proveedores
End Sub

Private Sub mnComContFactu_Click()
'Contabilizar Facturas
    AbrirListado (224) 'Para pedir datos
End Sub

Private Sub mnComDirecciones_Click()
    frmComDirecciones.Show vbModal
End Sub



Private Sub mnComEstAlbarxProve_Click()
'Listado de compras por proveedor
    AbrirListadoOfer (312)
End Sub

Private Sub mnComEstComprasxFam_Click()
'Listado de compras por Familia
    AbrirListadoOfer (311)
End Sub

Private Sub mnComEstComprasxProve_Click()
'Listado de compras por proveedor
    AbrirListadoOfer (310)
End Sub

Private Sub mnComEtiqProve_Click()
'Etiquetas de proveedores
     AbrirListadoOfer (305) '305: Informe Etiquetas de Proveedores
End Sub

Private Sub mnComFacturar_Click()
    frmComFacturar.Show vbModal
End Sub

Private Sub mnComHcoAlbaranes_Click()
'Historico albaranes de compras a proveedores
    frmComEntAlbaranes.EsHistorico = True
    frmComEntAlbaranes.Show vbModal
End Sub

Private Sub mnComHcoFacturas_Click()
    frmComHcoFacturas.hcoCodMovim = ""
    frmComHcoFacturas.Show vbModal
End Sub

Private Sub mnComHcoPedidos_Click()
    frmComEntPedidos.MostrarDatos = ""
    frmComEntPedidos.EsHistorico = True
    frmComEntPedidos.Show vbModal
End Sub

Private Sub mnComInfProve_Click()
'Informe de Proveedores
    AbrirListado (58)   ': Informe Proveedores
End Sub

Private Sub mnComPedMant_Click()
'Mnatenimiento de Pedidos de compras
    frmComEntPedidos.MostrarDatos = ""
    frmComEntPedidos.EsHistorico = False
    frmComEntPedidos.Show vbModal
End Sub


Private Sub mnComPreProve_Click(Index As Integer)
    Select Case Index
    Case 0
         frmComPreciosProv.Show vbModal
    Case 1
        frmComDtosFamMarca.Show vbModal
    Case 3
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene permiso", vbExclamation
        Else
            'Mostraremos el form para que pida la fecha
            frmComPreciosActualizar.Show vbModal
        End If
    End Select
    
End Sub

Private Sub mnComProveedores_Click()
'Compras. Mantenimiento de Proveedores
    frmComProveedores.Show vbModal
End Sub


Private Sub mnComProveVarios_Click()
'Proveedores varios
    frmComProveV.Show vbModal
End Sub

Private Sub mnComPteFacturar_Click()
'Listado de Albaranes pendientes de Factura
    AbrirListadoOfer (308) '308: List. Albaranes pte facturar
End Sub

Private Sub mnComPteRecibir_Click()
'Listado de material pendiente de recibir
    AbrirListadoOfer (307) '307: List. Materia pte recibir
End Sub

Private Sub mnConfManteUsuarios_Click()
'Mantenimiento de Usuarios

      frmMantenusu.Show vbModal
End Sub

Private Sub mnConfParamAplic_Click()
'Parametros de la Aplicación
    Screen.MousePointer = vbHourglass
    Load frmConfParamAplic
    frmConfParamAplic.Show vbModal
    
End Sub



Private Sub mnConfParamGenerales_Click()
    
'Parametros generales de la Empresa
    frmConfParamGral.Show vbModal
End Sub



Private Sub mnConfParamRpt_Click()
'Parametros para informes de Crystal Report
    frmConfParamRpt.Show vbModal
End Sub

Private Sub mnConTMovimiento_Click()
'Mantenimientos de los tipos de movimientos
    frmConfTipoMov.Show vbModal
End Sub


Private Sub mnCoupages_Click()
    frmAlmCoupage.DatosADevolverBusqueda2 = ""
    frmAlmCoupage.Show vbModal
End Sub

Private Sub mnCtasContab_Click()
    frmAlmContab.Show vbModal
End Sub

Private Sub mnEcoenves_Click()
    AbrirListado2 24
End Sub

Private Sub mnEliminarArticulos_Click()
    frmVarios.Opcion = 1
    frmVarios.Show vbModal
End Sub

Private Sub mnEliminarFacturas_Click()
    AbrirListado 97
End Sub

Private Sub mnEnvioFactuasMail_Click()
    AbrirListadoOfer 315
End Sub


Private Sub mnEtiqEstanteria_Click()
    AbrirListado 94
End Sub

Private Sub mnEtiqMante_Click()
    AbrirListado 79
End Sub

Private Sub mnEtiquetasArticulos_Click()
    mnEtiqEstanteria_Click
End Sub

Private Sub mnEtiquetasBultos_Click()
'Listado de etiquetas de los bultos
    AbrirListado 95
End Sub

Private Sub mnExportarFacturasPDF_Click()
     AbrirListadoOfer 316
End Sub

Private Sub mnFacActividades_Click()
    frmFacActividades.Show vbModal
End Sub

Private Sub mnFacAgentesCom_Click()
    frmFacAgentesCom.Show vbModal
End Sub

Private Sub mnFacAlbMostrador_Click()
'Abre el formulario de Albaranes para introducir el Albaran de Mostrador
'y desde este generar la Factura de mostrador
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALM"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub


Private Sub mnFacAlbRectifica_Click()
'Facturas Rectificativas
    'Abre el formulario de Albaranes para introducir el Albaran Rectificativo
    'y desde este generar la Factura Rectificativa
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ART"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacAlbxArtic_Click()
'Informe de Albaranes por Articulo
    AbrirListadoPed (49)
End Sub

Private Sub mnFacAltas_Click()
'Informe de Altas de Nuevos Clientes
    AbrirListadoOfer (48) '48: Informes Altas Clientes
End Sub

Private Sub mnFacBancosPropios_Click()
    frmFacBancosPropios.Show vbModal
End Sub

Private Sub mnFacBoniFact_Click()
'Bonificacines factura
    frmFacBonificacion.Show vbModal
End Sub

Private Sub mnFacCartaClien_Click()
'Cartas a clientes
     AbrirListadoOfer (91) '91: Informe Cartas a Clientes
End Sub

Private Sub mnFacCartas_Click()
'Mantenimiento de Cartas
    frmFacCartasOferta.Show vbModal
End Sub


Private Sub mnFacClientes_Click()
'Mantenimiento de Clientes
    frmFacClientes.Show vbModal
End Sub

Private Sub mnFacClientesV1_Click()
'Mantenimiento de Clientes Varios
    frmFacClientesV.Show vbModal
End Sub



Private Sub mnFacContFactu_Click()
'Contabilizar Facturas
    AbrirListado (223) 'Para pedir datos
End Sub

Private Sub mnFacDescuentos_Click()
    frmFacDtosFamMarca.Show vbModal
End Sub



Private Sub mnFacEntAlbaran_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALV"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub





Private Sub mnFacEstDetalleFac_Click()
'Detalle facturacion clientes
     AbrirListadoOfer (231)
End Sub

Private Sub mnFacEstMargenVtas_Click()
    'Estadistica margen ventas por artículo
        AbrirListado (246)
End Sub

Private Sub mnFacEstVentaAceite_Click()
    AbrirListado2 19
End Sub

Private Sub mnFacEstVentaCliente_Click()
'Estadistica Ventas por cliente
    AbrirListadoPed (227)
    BorrarTempInformes
End Sub

Private Sub mnFacEstVentaFam_Click()
'Listado de estadistica ventas por familia de articulo
    AbrirListadoOfer (230)
End Sub

Private Sub mnFacEstVentaMes_Click()
'Estadistica Ventas por Meses
    AbrirListadoPed (229)
    
End Sub

Private Sub mnFacEstVentaTraba_Click()
'Estadistica Ventas por Trabajador
    AbrirListadoPed (228)
End Sub

Private Sub mnFacEtiqClien_Click()
'Etiquetas de clientes
     AbrirListadoOfer (90) '90: Informe Etiquetas de Clientes
End Sub

Private Sub mnFacFacturarAlb_Click()
'Facturacion de Albaranes de Ventas
    frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (52)
End Sub

Private Sub mnFacFormasPago_Click()
    frmFacFormasPago.Show vbModal
End Sub



Private Sub mnFacHcoAlbaranes_Click()
'Histórico de Albaranes eliminados
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALV"
    frmFacEntAlbaranes.EsHistorico = True
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacHcoFacturas_Click()
'Histórico de Facturas
    frmFacHcoFacturas.hcoCodMovim = ""
    frmFacHcoFacturas.Show vbModal
End Sub




Private Sub mnFacInactivos_Click()
'Informe de Clientes Inactivos
    AbrirListadoOfer (46) '46: Informes Clientes Inactivos
End Sub

Private Sub mnFacIncidencias_Click()
    frmIncidencias.Show vbModal
End Sub

Private Sub mnFacIncumPlazos_Click()
'Incumplimiento de los Plazos de Entrega
    
    AbrirListadoPed (51)
End Sub

Private Sub mnFacInfClientes_Click()
'Informe de Clientes
    AbrirListadoOfer (47) '47: Informes Clientes
End Sub

Private Sub mnFacInfMargenes_Click()
'Informe control margenes de tarifas
    AbrirListado (245)
End Sub






Private Sub mnFacOfertas_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn
    
    Select Case Index
    Case 0, 5
            'Private Sub mnFacEntOfertas_Click()
        frmFacEntOfertas.DatosOferta = ""
        frmFacEntOfertas.EsHistorico = Index = 5
            frmFacEntOfertas.Show vbModal

    Case 1
            'Private Sub mnFacGrupoPlant_Click()
            'Mantenimiento de Grupos de Plantillas
        frmFacGrupoPlantilla.Show vbModal
    
    Case 2
            'Private Sub mnFacPlantillas_Click()
            'Mantenimiento de Plantillas
        frmFacPlantilla.Show vbModal
        
    Case 3
            ' Private Sub mnFacOfeEfectuadas_Click()
            'Listado de Ofertas Efectuadas
        AbrirListadoOfer (34) '34: Informe Ofertas Efectuadas
    
        
        
    'case 4  'Es la barra separadora
    
    Case 6
        
            'Private Sub mnFacTrasHist_Click()
            'Traspaso de Ofertas a las tablas de Historico
        frmListadoOfer.OpcionListado = 36
        AbrirListadoOfer (36) 'NO IMPRIME LISTADO, hace traspaso de Ofertas de la tabla (scapre) a (schpre)

    
    End Select
End Sub

Private Sub mnFacPedidos_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn
  
    Select Case Index
    Case 0, 1
        'Mantenimiento de Pedidos  Y Histórico de Pedidos
        frmFacEntPedidos.EsHistorico = Index = 1
        frmFacEntPedidos.Show vbModal
    
    'Case 2  es la barra de separacion
    
    Case 3
        'Confirmar pedido   mnFacConfirmPed_Click
        AbrirListadoOfer (40)
        
    Case 4
        'Pedido por articulo
        'Private Sub mnFacPedidoxArtic_Click()
        'Informe de Pedidos por Articulo
        AbrirListadoPed (41)
        
    Case 5
        'Private Sub mnFacPedidoxClien_Click()
        'Informe de Pedidos por Cliente
        AbrirListadoPed (44)
        
        
    Case 6
        'Private Sub mnFacDispStock_Click()
        'Resumen de Disponibilidad de Stocks
        AbrirListadoPed (42)
    Case 8
        frmFacConsultaPrecios.Show vbModal
    Case 9
        frmVarios.Opcion = 2
        frmVarios.Show vbModal
    End Select
End Sub



Private Sub mnFacPreEspecial_Click()
    frmFacPreciosEspecial.CadenaSituarData = ""
    frmFacPreciosEspecial.Show vbModal
End Sub

Private Sub mnFacPreFacturar_Click()
' Previsión Facturacion de Albaranes
    frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO
End Sub

Private Sub mnFacPromociones_Click()
    frmFacPromociones.Show vbModal
End Sub

Private Sub mnFacReImpFactu_Click()
'Reimprimir Factuas ya contabilizadas
    AbrirListadoOfer 226
End Sub

Private Sub mnFacRutas_Click()
    frmFacRutas.Show vbModal
End Sub

Private Sub mnFacSituaciones_Click()
    frmFacSituaciones.Show vbModal
End Sub

Private Sub mnFactActPrecios_Click()
'Actualizar precios actuales y especiales
    frmFacActPrecios.Show vbModal
End Sub

Private Sub mnFacTarifasPrecios_Click()
'Listado Precios
    frmFacTarifasPrecios.Show vbModal
End Sub

Private Sub mnFacTarifasVen_Click()
'Tarifas Venta
     frmFacTarifas.Show vbModal
End Sub





Private Sub mnFacturarPresupuestos_Click()
        frmListadoPed.CodClien = "ALZ" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
End Sub

Private Sub mnFacZonas_Click()
    frmFacZonas.Show vbModal
End Sub

Private Sub mnFacFormasEnvio_Click()
    frmFacFormasEnvio.Show vbModal
End Sub



Private Sub mnHcoMaten_Click()
 '   frmManMantenimientosAnu.Show vbModal
End Sub


Private Sub mnInfManteAnulados_Click()
    AbrirListado 76
End Sub



Private Sub mnInfTeoMant_Click()
    AbrirListado 77
End Sub



Private Sub mnInventarioAceite_Click()
    AbrirListado2 22
End Sub

'Private Sub mnListadoReparacionesEfectuadas_Click()
'    AbrirListado2 1
'End Sub

Private Sub mnManAltas_Click()
'Listado Altas de Mantenimientos
    AbrirListado 73
End Sub

Private Sub mnManEntrada_Click()
 '   frmManMantenimientos.Show vbModal
End Sub

Private Sub mnManFactAlb_Click()
'Facturacion de Mantenimientos
     AbrirListadoPed (75) 'NO IMPRIME LISTADO
End Sub

Private Sub mnManFichas_Click()
'Listado Fichas de Mantenimientos
    AbrirListado 72
End Sub

Private Sub mnManListado_Click()
'Listados de Mantenimientos
    AbrirListado 70
End Sub

Private Sub mnManPrevFac_Click()
' Previsión Facturacion de Albaranes de Mantenimiento
'    frmListadoPed.CodClien = "ALM" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (74) 'NO IMPRIME LISTADO
End Sub

Private Sub mnManRevisiones_Click()
'Listado Revisiones de Mantenimientos
     AbrirListado 71
End Sub

Private Sub mnManServicioAsisTecn_Click()
    frmManSat.Show vbModal
End Sub



Private Sub mnManteneLOG_Click()
    Screen.MousePointer = vbHourglass
    Load frmLog
    DoEvents
    frmLog.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnManTiposContrato_Click()
'    frmManTiposContrato.Show vbModal
End Sub


Private Sub mnNuevosPuntosMenuTraza_Click(Index As Integer)
    If Index = 1 Then
        'Control calidad
        frmProdCalidad.Show vbModal
    End If
        
End Sub

Private Sub mnOliTarifasOfertas_Click(Index As Integer)
    If Index = 1 Or Index = 2 Then
        frmOliTarifaOferta.EsTarifa = (Index = 2)
        frmOliTarifaOferta.Show vbModal
    ElseIf Index = 3 Then
                frmOliCrearTO1.vOpcion = 0
                frmOliCrearTO1.Show vbModal
    ElseIf Index = 4 Then
        frmOliCrearTO1.vOpcion = 1   'Copiar tarifa
        frmOliCrearTO1.Show vbModal
    ElseIf Index = 5 Then
        
        frmOliTarifaOferHco.Show vbModal
    End If
End Sub





Private Sub mnPartida2_Click(Index As Integer)
    If Index = 0 Then
        frmAlmPartidas.DatosADevolverBusqueda = ""
        frmAlmPartidas.Show vbModal
    End If
        
    'MAYO.  La smoval lotes la dejo aparte de momento
    
    'la produccion y el cupaje
    If Index = 1 Then
        frmAlmpartidasMov.VerPartida = 0
        frmAlmpartidasMov.Show vbModal   'igual entrando como root(es decir YO) podriam preguntar
    End If
    
    'Este lee de smovallotes  MAYO 2010. Lo comento y qu
    'Ma adelante, cuando este todo el proceso instaurado.. ESTE sera el bueno. De moemnto 1 saco de cemnto
    'If Index = 1 Then frmAlmpartidasMovNuevo.Show vbModal
End Sub

Private Sub mnPreciosTarifasCorreccion_Click()
    AbrirListado 247
End Sub





Private Sub mnproduccion1_1_Click(Index As Integer)
    Select Case Index
    Case 0
        If vParamAplic.QUE_EMPRESA = 4 Then
            frmProdDepositosVall.Show vbModal
            
        Else
            frmProdDepositos.Show vbModal
        End If
    
    Case 1
    
        frmProduVarios.Opcion = 2
        frmProduVarios.Show vbModal
    
    
    Case 2
        'FILTRADO
        frmProduVarios.Opcion = 4
        frmProduVarios.Show vbModal
    
    Case 3
        'VACIADO
        frmProduVarios.Opcion = 3
        frmProduVarios.Show vbModal
    
    End Select
End Sub

Private Sub mnproduccion1_Click(Index As Integer)
Dim Permis As String
Dim I As Integer

    '- Las líneas de producción y la pistola sólo las tiene que poder ver José y Yo.
    '- Autorizar a todos a poder ver: lote de trazabilidad, mantenimiento de palets y mantenimiento de etiquetas.

    If Index >= 3 Then
        
        If vUsu.Nivel > 2 Then
            MsgBox "No tiene permiso", vbExclamation
            Exit Sub
        End If
        
        
  
    End If
    
    Select Case Index
    Case 0
        frmProdOrden.DatosADevolverBusqueda2 = ""
        frmProdOrden.Show vbModal
    Case 1
        frmAlmDescCostesTasas.Show vbModal
    Case 3
    '       Julio 2014
    '    Si tiene el punto visible, tiene permiso
    '    frmProdNUEVA.PermisoPlanning = Mid(Permis, 2, 1) = "1"
    '    frmProdNUEVA.PermisoProduccion = Mid(Permis, 1, 1) = "1"
    
        frmProdNUEVA.PermisoPlanning = True
        frmProdNUEVA.PermisoProduccion = True
        frmProdNUEVA.Show vbModal
    Case 4
       ' frmProdSeleccionarLineaProd.YaProducidas = True
       ' frmProdSeleccionarLineaProd.Show vbModal
        MsgBox "Falta"
    Case 5
        
        frmProdNueTraza2.QueTrazabilidad = 0
        frmProdNueTraza2.Show vbModal

        
    Case 6
       
        
        
    Case 8
        frmProdPalets.DatosADevolverBusqueda2 = ""
        frmProdPalets.Show vbModal
    
    Case 9
        frmProdEtiquetas.Show vbModal
    Case 11
        frmPist1.Show vbModal
    End Select
End Sub




'Private Function SumaUnosPermisos(Permiso As String) As Byte
'Dim i As Integer
'    SumaUnosPermisos = 0
'    For i = 1 To Len(Permiso)
'        If Mid(Permiso, i, 1) = "1" Then SumaUnosPermisos = SumaUnosPermisos + 1
'    Next
'
'
'    'SumaUnosPermisos = 1
'End Function

Private Sub mnRecalculoPRecioSt_Click(Index As Integer)
    'Recalcular precios standard a partir del precio standard de
    ' la materia prima
    AbrirListado2 30
End Sub

Private Sub mnRecupFac_Click()
'recuperar facturas
    'abrimos albaranes de mostrador
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALM"
    'le indicamos q estamos recuperando facturas
    frmFacEntAlbaranes.RecuperarFactu = True
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnRegistros_Click(Index As Integer)
Dim Ind As Integer
    Ind = Index
    Select Case Ind
    Case 1
        'limpieza
        frmRegistros.Show vbModal
    
    Case 2
        'Mto preventivo
        frmRegistrosManPrev.Show vbModal
        
    Case 3
        frmRegLisRev.Show vbModal
        
        
    Case 4
        frmVarios.Opcion = 7
        frmVarios.Show vbModal
        
    Case 5
        frmRegRevision.Show vbModal
    Case 6
        frmRegAccCorrec.Show vbModal
    Case 8
        ComprobarDocumentosPendientes False
    End Select
End Sub





Private Sub mnRevisarMultibase_Click()
    AbrirListado2 3
End Sub

Private Sub mnRutas_Click(Index As Integer)
    If Index = 0 Then
        frmFacRutasEnvio.Show vbModal
    Else
        frmFacVehiculos.Show vbModal
    End If
End Sub

'Private Sub mnPedirPwd_Click()
'Dim Anterior As Boolean
'
'    Anterior = Me.mnPedirPwd.Checked
'    vConfig.PedirPasswd = Not Anterior
'    If vConfig.Grabar = 1 Then
'        Me.mnPedirPwd.Checked = Anterior
'    Else
'        Me.mnPedirPwd.Checked = Not Anterior
'    End If
'End Sub


Private Sub mnSeleccionarImpresora_Click()
    Screen.MousePointer = vbHourglass
    Me.CommonDialog1.ShowPrinter
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnServicios_Click(Index As Integer)
    If Index = 0 Then Exit Sub  'La barra no puede
    Select Case Index
    Case 1
        frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes.hcoCodTipoM = "ALS"
        frmFacEntAlbaranes.EsHistorico = False
        frmFacEntAlbaranes.RecuperarFactu = False
        frmFacEntAlbaranes.Show vbModal
    Case 2
        frmListadoPed.CodClien = "ALS" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
    Case 4
        frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes.hcoCodTipoM = "ALI"
        frmFacEntAlbaranes.EsHistorico = False
        frmFacEntAlbaranes.RecuperarFactu = False
        frmFacEntAlbaranes.Show vbModal
    Case 5
        frmListadoPed.CodClien = "ALI" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
    End Select
End Sub

Private Sub mnSociosProveedores_Click(Index As Integer)
    Select Case Index
    Case 0
        'Cambiar precios proveedores /socios
         AbrirListado2 7
         
    Case 1
        'Liquidacion SOCIOS
        AbrirListado2 8
        
    Case 2
        'Impresion facturas proveedores
        AbrirListado2 9
        
    Case 3
        'MsgBox "En desarrollo", vbExclamation
        'Asociar albaranes compras / vetnas
         frmComprasVentas.Show vbModal
    
    Case 4
        'listado trazabilidad
        AbrirListado2 15
    End Select
End Sub

Private Sub mnSoporte_Click(Index As Integer)

    Select Case Index
    Case 4
       
        Screen.MousePointer = vbHourglass
        LanzaHome ("websoporte")
        Screen.MousePointer = vbDefault
    
    Case 7
        frmVarios.Opcion = 8
        frmVarios.Show vbModal
    Case 9
        'Acerca de
        Screen.MousePointer = vbDefault
        frmMensajes.OpcionMensaje = 3
        frmMensajes.Show vbModal
    End Select
    
End Sub



Private Sub mnTicket_Click(Index As Integer)
    
    If Index > 0 Then AbrirListado2 12 + Index

    
End Sub

Private Sub mnTiposArticulos_Click()
    frmAlmTipoArticulo.Show vbModal
End Sub

Private Sub mnSalir_Click()
    End
End Sub







Private Sub mnTraspasoFraAVAB_Click(Index As Integer)
    If Index = 1 Then
        'MsgBox "En desarrollo", vbExclamation
        frmFacTraspasoAVAB.Show vbModal
    End If
End Sub

Private Sub mnTraspasoMante_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.OpcionMensaje = 18
    frmMensajes.Show vbModal
End Sub










Private Sub mnTrazaNueva_Click(Index As Integer)
    

    Select Case Index
    Case 0
        
            frmProdTrazaVer2.Show vbModal
    Case 1
            frmFacTrazabilidad3.Opcion = 0
            frmFacTrazabilidad3.Show vbModal
    Case 3
    
            frmFacTrazabilidad3.Opcion = 1
            frmFacTrazabilidad3.Show vbModal
    Case 4
            frmFacTrazabilidad.Show vbModal
    
    Case 5
         frmFacTrazabilidad2.Show vbModal
     
     End Select
        
End Sub

Private Sub mnUtiBuscarErrConCli_Click()
'Facturas pendientes de contabilizar (CLIENTES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 6
    frmUtilidades.Show vbModal
End Sub

Private Sub mnUtiBuscarErrConPro_Click()
'Facturas pendientes de contabilizar (PROVEEDORES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 7
    frmUtilidades.Show vbModal
End Sub


Private Sub mnUtiBuscarErrFac_Click()
'Buscar errores en nº de factura (solo en facturas de clientes)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 5
    frmUtilidades.Show vbModal
End Sub





Private Sub mnUtiDeclaracionLOM_Click()
'Declaracion LOM
    frmUtDeclara.Show vbModal
End Sub

'Private Sub mnUtiMensNuevo_Click()
''Nuevo mensaje en la utilidad de mensajeria interna
'    frmMensaje2.Show vbModal
'End Sub
'
'Private Sub mnUtiMensTipMen_Click()
'    frmTiposMensajes.Show vbModal
'End Sub



Private Sub mnUtiUsuActivos_Click()
'Muestra si hay otros usuarios conectados a la Gestion
Dim SQL As String
Dim I As Integer


    On Error GoTo eUsacti

    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    If CadenaDesdeOtroForm <> "" Then
        I = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            SQL = RecuperaValor(CadenaDesdeOtroForm, I)
            If SQL <> "" Then Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
            I = I + 1
        Loop Until SQL = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningun usuario, además de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    
eUsacti:
    If Err.Number <> 0 Then
        CadenaDesdeOtroForm = "Error obteniendo PCs conectados " & vbCrLf
        MsgBox CadenaDesdeOtroForm, vbExclamation
        Err.Clear
    End If
    CadenaDesdeOtroForm = ""
End Sub



Private Sub mnVentasAgente_Click()
    AbrirListado2 20
End Sub

Private Sub mnVentasPorProveedor_Click()
    AbrirListado2 6
End Sub

Private Sub mnVerAvisos_Click()
    If TieneAvisosPendientes Then
        frmAlertas.Show vbModal
    Else
        MsgBox "No hay avisos para mostrar", vbInformation
    End If
End Sub







Private Sub mnVtasAgrupadox_Click()
    frmListado2.Opcion = 26
    frmListado2.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    
    Select Case Button.Index
    Case 1 'Mantenimiento de Artículos
        mnAlmArticulosMto_Click 0
    Case 2 'Movimientos Articulos
        mnAlmMovimArticulos_Click
        
    Case 5 'Mantenimiento Clientes
        mnFacClientes_Click
    Case 6 'Mantenimiento Proveedores
        mnComProveedores_Click
        
    Case 9 'Ofertas a Clientes
        mnFacOfertas_Click 0
    Case 10 'Pedidos a Clientes
        'mnFacEntPedidos_Click
        mnFacPedidos_Click 0
    Case 11 'Albaranes a Clientes
        If vUsu.TrabajadorB Then
            mnAlbaranesB_Click
        Else
            mnFacEntAlbaran_Click
        End If
    Case 12 'Hist. Albaranes (Facturas)
        mnFacHcoFacturas_Click
        
    Case 15 'Pedidos de Proveedores
        mnComPedMant_Click
    Case 16 'Albaranes de Proveedores
        mnComAlbMan_Click
    Case 17 'Facturas de Proveedores
        mnComHcoFacturas_Click
    Case 18 'Recepcion Fact. Prove
        If Me.mnComFacturar.visible And Me.mnComFacturar.Enabled Then mnComFacturar_Click
        
    Case 21 'Mantenimientos
        mnManEntrada_Click
    Case 22 'Nº Serie
        'mnRepNumSerie_Click
        
    Case 24 'Gastos Técnicos
        mnAdmGastosTec_Click
    Case 25
        'Consulta precio articulo
        mnFacPedidos_Click 8
        
    Case 26 'Entrada al TPV
        'mnTPVpantallaVenta_Click
    Case 27
        'cambiar empresa
        mnCambioEmpresa_Click
        
    Case 28

        mnAgenda_Click
        
    Case 30
        frmMoixent.Show vbModal
    Case 32 'Salir
        mnSalir_Click
    End Select
End Sub


Private Sub PonerDatosVisiblesForm()
'Escribe texto de la barra de la aplicación
Dim cad As String
    cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    cad = cad & ", " & Format(Now, "d")
    cad = cad & " de " & Format(Now, "mmmm")
    cad = cad & " de " & Format(Now, "yyyy")
    cad = "    " & cad & "    "
    Me.StatusBar1.Panels(5).Text = cad
    If vEmpresa Is Nothing Then
        Caption = "ARIGES" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
        'Panel con el nombre de la empresa
        Me.StatusBar1.Panels(2).Text = "Falta configurar"
    Else
        Caption = "Gestión" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vEmpresa.nomempre & "  -    Usuario: " & vUsu.Nombre
        If vUsu.TrabajadorB Then Caption = Caption & "   ****"
        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
    End If
End Sub


Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim cad As String

    
    For Each T In Me
        cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            If LCase(Mid(T.Caption, 1, 1)) <> "-" Then T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnConfParamAplic = True
    Me.mnConfParamGenerales = True

    Me.mnSalir.Enabled = True
    Me.mnCambioEmpresa.Enabled = True
End Sub

'-------------------------------------
'Pondremos todos los que menus a visibles. Y luego ya , en f
Private Sub ReestablecerMenus()
Dim T
      For Each T In Me
        If Mid(T.Name, 1, 2) = "mn" Then T.visible = True
    Next
End Sub

Private Sub PonerMenusNivelUsuario()
Dim b As Boolean

    b = (vUsu.Nivel = 0) Or (vUsu.Nivel = 1)  'Administradores y root

    Me.mnConfParamAplic = b
    mnConfManteUsuarios = b
    
    mnUsuarios.Enabled = b
    mnNuevaEmpresa.Enabled = b
    mnPedirPwd.Enabled = b
    'Me.mnUtiConnActivas.Enabled = (vUsu.Nivel = 0) 'solo para root
    

    b = vUsu.Nivel = 3  'Es usuario de consultas
    If b Then
        'Inventario
        Me.mnAlmTomaInven.Enabled = False
        Me.mnAlmEntradaInve.Enabled = False
        Me.mnAlmActualizarInve.Enabled = False
        Me.mnAlmListadoInve.Enabled = False
        Me.mnAlmValoracionInve.Enabled = False
        'Antes
        'Me.mnFacTrasHist.Enabled = False
        mnFacOfertas(6).Enabled = False
        
        
        'Facturacion Ventas
        Me.mnFacFacturarAlb.Enabled = False
        Me.mnFacContFactu.Enabled = False
        
        'Facturacion Compras
        Me.mnComFacturar.Enabled = False
        Me.mnComContFactu.Enabled = False
        
        'Reparaciones
        'Me.mnRepFactAlb.Enabled = False
        
        'Mantenimientos
        Me.mnManFactAlb.Enabled = False
    End If
End Sub



Private Sub LanzaHome(Opcion As String)
Dim I As Integer
Dim cad As String

    On Error GoTo ELanzaHome

'    LanzaHome = False
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", Opcion, "codigo", "1", "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en Parámetros de la Aplicación.", vbExclamation
        Exit Sub
    End If

    If Opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision


'    I = FreeFile
'    cad = ""
'    Open App.Path & "\lanzaexp.dat" For Input As #I
'    Line Input #I, cad
'    Close #I

    'Lanzamos
    If LanzaHomeGnral(CadenaDesdeOtroForm) Then Espera 2
    
'    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'    If vConfig.Explorador <> "" Then
'        Shell vConfig.Explorador & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'        LanzaHome = True
'    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub



Private Sub LeerEditorMenus()
Dim SQL As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from usuarios.appmenus where aplicacion='Arioli'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from usuarios.appmenususuario where aplicacion='Ariges" & vEmpresa.codempre & "' and codusu = " & Val(Right(CStr(vUsu.Codigo), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3)
            If Right(miRsAux.Fields(3), 1) <> "|" Then SQL = SQL & "|"
            SQL = SQL & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "·" & SQL
        For Each T In Me.Controls
            If TypeOf T Is Menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                'Debug.Print C
                If InStr(1, SQL, C) > 0 Then
                    
                    'Stop
                    T.visible = False
                End If
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index & "|"
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function



Private Sub PoneBarraMenus()
'Para cada boton de la toolbar comprobar que el menu con el que se corresponde
'esta visible y activado, y ponerle los mismos valore que tenga el menu
Dim Activado As Boolean
Dim b As Boolean

    On Error GoTo 0
    
    '-----------------------------------------------------------
    'Articulos
    Me.Toolbar1.Buttons(1).visible = ComprobarBotonMenuVisible(Me.mnAlmArticulos, Activado)
    Me.Toolbar1.Buttons(1).Enabled = Activado

    'Movimientos de Articulos
    Me.Toolbar1.Buttons(2).visible = ComprobarBotonMenuVisible(Me.mnAlmMovimArticulos, Activado)
    Me.Toolbar1.Buttons(2).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Clientes
    Me.Toolbar1.Buttons(5).visible = ComprobarBotonMenuVisible(Me.mnFacClientes, Activado)
    Me.Toolbar1.Buttons(5).Enabled = Activado
    
    'Proveedores
    Me.Toolbar1.Buttons(6).visible = ComprobarBotonMenuVisible(Me.mnComProveedores, Activado)
    Me.Toolbar1.Buttons(6).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Ofertas Clientes
    Me.Toolbar1.Buttons(9).visible = ComprobarBotonMenuVisible(Me.mnFacOfertas(0), Activado)
    Me.Toolbar1.Buttons(9).Enabled = Activado
    
    'Pedidos Clientes
    Me.Toolbar1.Buttons(10).visible = ComprobarBotonMenuVisible(mnFacPedidos(0), Activado)
    Me.Toolbar1.Buttons(10).Enabled = Activado
    
    'Albaranes Clientes
    Me.Toolbar1.Buttons(11).visible = ComprobarBotonMenuVisible(Me.mnFacEntAlbaran, Activado)
    Me.Toolbar1.Buttons(11).Enabled = Activado
    
    'Facturas Clientes
    Me.Toolbar1.Buttons(12).visible = ComprobarBotonMenuVisible(Me.mnFacHcoFacturas, Activado)
    Me.Toolbar1.Buttons(12).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Pedidos Proveedor
    'Comprobar que los menus del que cuelga no este bloqueado o invisible
    Me.Toolbar1.Buttons(15).visible = ComprobarBotonMenuVisible(Me.mnComPedMant, Activado)
    Me.Toolbar1.Buttons(15).Enabled = Activado

    'Albaranes Proveedor
    Me.Toolbar1.Buttons(16).visible = ComprobarBotonMenuVisible(Me.mnComAlbMan, Activado)
    Me.Toolbar1.Buttons(16).Enabled = Activado
    
    'Facturas Proveedor
    Me.Toolbar1.Buttons(17).visible = ComprobarBotonMenuVisible(Me.mnComHcoFacturas, Activado)
    Me.Toolbar1.Buttons(17).Enabled = Activado
    
    'Recepcion facturas de compras
    Me.Toolbar1.Buttons(18).visible = ComprobarBotonMenuVisible(Me.mnComFacturar, Activado)
    Me.Toolbar1.Buttons(18).Enabled = Activado


    '-----------------------------------------------------------
    'Mantenimientos
    b = False

    'Me.Toolbar1.Buttons(21).visible = ComprobarBotonMenuVisible(B, Activado)
    'Me.Toolbar1.Buttons(21).Enabled = Activado
    b = vParamAplic.Mantenimientos
    If b Then b = ComprobarBotonMenuVisible(mnManEntrada, Activado)
        
    Me.Toolbar1.Buttons(21).visible = b
    Me.Toolbar1.Buttons(21).Enabled = Activado
     
    
    'Nº Serie
    Me.Toolbar1.Buttons(22).visible = vParamAplic.Reparaciones
    Me.Toolbar1.Buttons(22).Enabled = False
    
    
    '-----------------------------------------------------------
    'Conuslta de precio
    Me.Toolbar1.Buttons(24).visible = ComprobarBotonMenuVisible(Me.mnFacPedidos(8), Activado)
    Me.Toolbar1.Buttons(24).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Gastos tecnicos
    Me.Toolbar1.Buttons(25).visible = ComprobarBotonMenuVisible(Me.mnAdmGastosTec, Activado)
    Me.Toolbar1.Buttons(25).Enabled = Activado
    
    
    'Nuevos botones
    'TPV
    'Me.Toolbar1.Buttons(26).visible = ComprobarBotonMenuVisible(mnTPVpantallaVenta, Activado)
    Me.Toolbar1.Buttons(26).visible = False
    
    'Cambio empresa. SIEMPRE estara visible
    'Me.Toolbar1.Buttons(27).visible = ComprobarBotonMenuVisible(mnCambioEmpresa, Activado)
    'Me.Toolbar1.Buttons(27).Enabled = Activado
    Me.Toolbar1.Buttons(27).visible = True
    Me.Toolbar1.Buttons(27).Enabled = True
    
    'Agenda
    Me.Toolbar1.Buttons(28).visible = ComprobarBotonMenuVisible(mnAgenda, Activado)
    Me.Toolbar1.Buttons(28).Enabled = Activado
    
    
    'Maenteinimiento de entradas MOIXENT
    Me.Toolbar1.Buttons(30).visible = ComprobarBotonMenuVisible(Me.mnMoixentMov(1), Activado)
    Me.Toolbar1.Buttons(30).Enabled = Activado
    
End Sub




Private Function ComprobarBotonMenuVisible(objMenu As Menu, Activado As Boolean) As Boolean
'Comprueba si el icono de la barra se debe activar/desactivar o si se debe poner
'visible o invisible. Para ello comprueba si su correspondiente entrada de menu
'esta activada/desactiva o visible/invisible
'(se comprueba hasta q se encuentra el false o se llega al padre)
Dim nomMenu As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cad As String
Dim b As Boolean


    On Error GoTo EComprobar
    
    b = objMenu.visible
    Activado = objMenu.Enabled
    
    If b = False Then
        ComprobarBotonMenuVisible = b
    Else
    
        nomMenu = objMenu.Name
        
        Set RS = New ADODB.Recordset
        
        'Obtener el padre del menu
        SQL = "select padre from usuarios.appmenus where aplicacion='Arioli' and name=" & DBSet(nomMenu, "T")
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            cad = RS.Fields(0).Value
        End If
        RS.Close
        
        b = True
        While b And cad <> ""
                SQL = "Select name,padre from usuarios.appmenus where aplicacion='Arioli' and contador= " & cad
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    cad = RS!Padre
                    nomMenu = RS!Name
                End If
                RS.Close
                
                'comprobar si el padre esta bloqueado
                SQL = "Select count(*) from usuarios.appmenususuario where aplicacion='Ariges" & vEmpresa.codempre & "' and codusu=" & Val(Right(CStr(vUsu.Codigo), 3))
                SQL = SQL & " and tag='" & nomMenu & "|'"
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS.Fields(0).Value > 0 Then
                    'Esta bloqueado el menu para el usuario
                    b = False
                End If
                RS.Close
                If cad = "0" Then cad = "" 'terminar si llegamos a la raiz
        Wend
        ComprobarBotonMenuVisible = b
        Set RS = Nothing
    End If
    
EComprobar:
    If Err.Number <> 0 Then Err.Clear
End Function



Private Sub AbrirListado2(KOpcion As Integer)
    Screen.MousePointer = vbHourglass
    frmListado2.Opcion = KOpcion
    frmListado2.Show vbModal
    Screen.MousePointer = vbDefault
End Sub







'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'
'
'   ARIDOC.  para los datos de ARIDOC reutilizare la conneion conta
'           con lo cual la cerrare y abrire tantas veces necesite
'


Private Sub HacerMenuARidoc(Opcion As Byte)
    
    If Conexion_Aridoc_(True) Then
        Select Case Opcion
        Case 0
            frmAridocConfig.Show vbModal
        End Select
    End If
    Conexion_Aridoc_ False
End Sub










Private Sub ComprobarDocumentosPendientes(DesdeArranqueAplicacion As Boolean)
Dim F As Date
Dim cad As String
Dim Aux As String
Dim J As Integer

    'DE MOMENTO SOLO PARA MORALES
    If vParamAplic.EsAVAB Then Exit Sub
    If Not vParamAplic.Produccion Then Exit Sub
    'Si no es visible no cpruebo nada
    If Not mnRegPpañ.visible Then Exit Sub
    
    
    F = Format(Now, "dd/mm/yyyy")
    If DesdeArranqueAplicacion Then
        AvisosFechaArranque True, F
    Else
        F = DateAdd("d", -1, F)
    End If
    If F >= Format(Now, "dd/mm/yyyy") Then Exit Sub
    
    Set miRsAux = New ADODB.Recordset
    
    'Comprobar documentos
    'Aqui tenemos los documentos de LIMPIEZA
    '                                                   k tenga definida la perioricidad
    cad = "select * from sregistros WHERE diasaviso > 0  and perioricidad >0"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        F = miRsAux!PrimeraFecha
        If Not IsNull(miRsAux!UltimoRealizado) Then F = miRsAux!UltimoRealizado
        Select Case CInt(miRsAux!Perioricidad)
            Case 2
                'Semanas
                CadenaDesdeOtroForm = "ww"
            Case 3
                'MEs
                CadenaDesdeOtroForm = "m"
            Case 4
                CadenaDesdeOtroForm = "yyyy"
            Case Else
                'dias
                CadenaDesdeOtroForm = "d"
        End Select
            F = DateAdd(CadenaDesdeOtroForm, CInt(miRsAux!NumPeriodo), F)
            'Dias de aviso
            F = DateAdd("d", -1 * CInt(miRsAux!diasaviso), F)
            If F <= Now Then cad = cad & Format(F, "dd/mm/yyyy") & "  " & miRsAux!Descripcion & " (" & miRsAux!idRegistro & ")" & vbCrLf
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If cad <> "" Then cad = "LIMPIEZA" & vbCrLf & cad
    
    
    'Mantenimiento PREVENTIVO
    Aux = "select * from sregistrosmanprev WHERE diasaviso > 0  and perioricidad >0"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not miRsAux.EOF
        F = miRsAux!PrimeraFecha
        If Not IsNull(miRsAux!UltimoRealizado) Then F = miRsAux!UltimoRealizado
        Select Case CInt(miRsAux!Perioricidad)
            Case 2
                'Semanas
                CadenaDesdeOtroForm = "ww"
            Case 3
                'MEs
                CadenaDesdeOtroForm = "m"
            Case 4
                CadenaDesdeOtroForm = "yyyy"
            Case Else
                'dias
                CadenaDesdeOtroForm = "d"
        End Select
            F = DateAdd(CadenaDesdeOtroForm, CInt(miRsAux!NumPeriodo), F)
            'Dias de aviso
            F = DateAdd("d", -1 * CInt(miRsAux!diasaviso), F)
            If F <= Now Then Aux = Aux & Format(F, "dd/mm/yyyy") & "  " & miRsAux!Denominacion & " (" & miRsAux!Codigo & ")" & vbCrLf
                
            
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Aux <> "" Then Aux = vbCrLf & "MANTENIMIENTO PREVENTIVO: " & vbCrLf & Aux
    cad = cad & Aux
    
    
    'CLORO
    Aux = DevuelveDesdeBD(conAri, "max(fech2)", "sregcloro", "1", 1)
    If Aux = "" Then
        'Aux = "Primer analisis no definido"
    Else
        F = CDate(Aux)
        J = DateDiff("d", Now, F)
        If J <= 3 Then
            Aux = "Ultima hoja analitica cloro: " & F
        Else
            Aux = ""
        End If
    End If
    
    
    If Aux <> "" Then Aux = vbCrLf & "CLORO: " & vbCrLf & Aux
    cad = cad & Aux
    
    
    
    
    
    
    Set miRsAux = Nothing
    CadenaDesdeOtroForm = ""
    AvisosFechaArranque False, F
    
    If cad <> "" Then
        cad = "FECHA    REGISTRO " & vbCrLf & String(30, "=") & vbCrLf & cad & vbCrLf
        MsgBox cad, vbInformation
    End If
End Sub




Private Sub AvisosFechaArranque(Leer As Boolean, ByRef Fecha As Date)
Dim NF As Integer
Dim C As String
    On Error GoTo EComprobarFechaArranque
    Fecha = Format(Now, "dd/mm/yyyy")
    C = App.Path & "\Revavi.dat"
    If Leer Then
        If Dir(C, vbArchive) <> "" Then
            NF = FreeFile
            Open C For Input As #NF
            Line Input #NF, C
            Close #NF
            If C <> "" Then
                If EsFechaOK(C) Then Fecha = Format(C, "dd/mm/yyyy")
            
            End If
        Else
            Fecha = DateAdd("d", -1, Fecha)
        End If
    Else
        'Guardar
        NF = FreeFile
        Open C For Output As #NF
        Print #NF, Format(Now, "dd/mm/yyyy")
        Close #NF

    End If

    Exit Sub
EComprobarFechaArranque:
    MuestraError Err.Number, "Comprobar fecha avisos documentos"
    
End Sub







































'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
' Todo el trozo de REPARCIONES esta comentado aqui aajo
'  NO quitar. Por si algun dia, en el futuro....
'###################################################################################
'###################################################################################
'###################################################################################
'###################################################################################
'
'Private Sub mnEstadisticaReparacionTecnico_Click()
'    AbrirListado2 2
'End Sub
'
'Private Sub mnUtiConnActivas_Click()
''ver las conexiones a donde apuntan
'Dim cad As String
'    cad = "Conexiones:" & vbCrLf
'    cad = cad & "------------------" & vbCrLf & vbCrLf
'    cad = cad & "Ariges: " & vbCrLf & Conn.ConnectionString & vbCrLf & vbCrLf
'    cad = cad & "Conta: " & vbCrLf & ConnConta.ConnectionString & vbCrLf
'    MsgBox cad, vbInformation
'End Sub
'
'Private Sub mnTrabaRealiz_Click()
'    frmManTraReali.Show vbModal
'End Sub
'
'
'
'Private Sub mnTiposAveria_Click()
'    frmtipave.Show vbModal
'End Sub
'
'Private Sub mnTPVcierreCaja_Click()
''Abre el informe de cierre de caja del dia en el TPV
'    AbrirListadoOfer (240)
'End Sub
'
'Private Sub mnTPVpantallaVenta_Click()
''Pantalla venta del TPV
'Dim Nom As String
'
'    'Antes de abrir la pantalla de venta comprobamos que podemos leer el terminal
'    'nom = ComputerNameTServer
'
'    Nom = ComputerName 'Nombre PC conectado por Terminal Server / local
'
'    If Trim(Nom) <> "" Then
'        frmFacTPVEnt.NomrePC_conectado = Nom
'        frmFacTPVEnt.Show
'    Else
''        'Terminal con el que trabajaremos, leemos el nombre del ordenador en local
''        'si no trabajamos en terminal server
''        nom = ComputerName
''        If Trim(nom) <> "" Then
''            frmFacTPVEnt.NomrePC_conectado = nom
''            frmFacTPVEnt.Show
''        Else
'            MsgBox "No se puedo establecer un terminal.", vbExclamation
''        End If
'    End If
'End Sub
'
'
'Private Sub mnTPVParamGen_Click()
'    'Parámetros generales del TPV
'    frmFacTPVParamG.Show vbModal
'End Sub
'
'Private Sub mnTPVParamTer_Click()
'    'parametros de los terminales(equipos) del TPV
'    frmFacTPVParamT.Show vbModal
'End Sub
'
'
'Private Sub mnTelefonia1_Click(Index As Integer)
'    Select Case Index
'    Case 0
'        'Matenimientos entradas telefonia
'        ' frmTelefonia.Show vbModal
'    Case 1
'        'Informes recargas
'        AbrirListado2 4
'    Case 2
'        'Faturar recargas
'        AbrirListado2 5
'    End Select
'End Sub
'
'Private Sub mnRepPrevFact_Click()
'' Previsión Facturacion de Albaranes de Reparacion
'    frmListadoPed.CodClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
'    AbrirListadoPed (50) 'NO IMPRIME LISTADO
'
'End Sub
'
'
'
'
'Private Sub mnRepAlbaranes_Click()
'    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
'    frmFacEntAlbaranes.hcoCodTipoM = "ALR"
'    frmFacEntAlbaranes.RecuperarFactu = False
'    frmFacEntAlbaranes.Show vbModal
'End Sub
'
'Private Sub mnRepAvisos_Click()
''Avisos de averias de clientes
'    frmRepAvisos.Show vbModal
'End Sub
'
'Private Sub mnRepControlRep_Click()
''Control de Reparaciones (para los Tecnicos)
'    frmRepEntReparaciones.EntradaEquipo = ""
'    frmRepEntReparaciones.ControlRep = True
'    frmRepEntReparaciones.EsHistorico = False
'    frmRepEntReparaciones.Show vbModal
'End Sub
'
'Private Sub mnRepEntReparacion_Click()
''Mantenimiento de Reparaciones
'    frmRepEntReparaciones.EntradaEquipo = ""
'    frmRepEntReparaciones.ControlRep = False
'    frmRepEntReparaciones.EsHistorico = False
'    frmRepEntReparaciones.Show vbModal
'End Sub
'
'Private Sub mnRepFactAlb_Click()
''Facturacion de Albaranes de Reparacion
'    frmListadoPed.CodClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
'    AbrirListadoPed (52)
'End Sub
'
'Private Sub mnRepHistorico_Click()
''Historico de las reparaciones
'    frmRepEntReparaciones.EntradaEquipo = ""
'    frmRepEntReparaciones.ControlRep = False
'    frmRepEntReparaciones.EsHistorico = True
'    frmRepEntReparaciones.Show vbModal
'End Sub
'
'
'Private Sub mnRepListAvisosPtes_Click()
''Listado de avisos de averias de clientes pendientes
'    AbrirListado (409)
'End Sub
'
'Private Sub mnRepListFrecuen_Click()
''Listado de Frecuencia de Reparaciones
'    AbrirListado (406)
'End Sub
'
'Private Sub mnRepListRepxClien_Click()
''Listado de las Reparaciones por cliente
'    AbrirListado (64)
'End Sub
'
'Private Sub mnRepListRepxDia_Click()
''Listado de las Reparaciones del dia
'    AbrirListado (63)
'End Sub
'
'Private Sub mnRepMotivosBaja_Click()
''Motivos baja equipos
'    frmRepMotivosBaja.Show vbModal
'End Sub
'
'Private Sub mnRepMotivosPend_Click()
''Motivos Pendientes Reparar
'    frmRepMotivosPend.Show vbModal
'End Sub
'
'Private Sub mnRepNumSerie_Click()
''Mantenimiento de Nºs de Serie
'    frmRepNumSerie.Show vbModal
'End Sub
'
'
