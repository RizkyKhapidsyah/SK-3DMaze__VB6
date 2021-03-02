VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Maze Maker"
   ClientHeight    =   7530
   ClientLeft      =   1335
   ClientTop       =   2085
   ClientWidth     =   9945
   Icon            =   "Maker.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7530
   ScaleWidth      =   9945
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7560
      TabIndex        =   410
      Top             =   1560
      Width           =   2295
   End
   Begin VB.FileListBox File 
      Height          =   675
      Left            =   7560
      Pattern         =   "*.map"
      TabIndex        =   408
      Top             =   2280
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "West"
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   406
      Top             =   4920
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "South"
      Height          =   255
      Index           =   3
      Left            =   8040
      TabIndex        =   405
      Top             =   4680
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "East"
      Height          =   255
      Index           =   2
      Left            =   8040
      TabIndex        =   404
      Top             =   4440
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "North"
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   403
      Top             =   4200
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   401
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   7560
      TabIndex        =   400
      Top             =   960
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   399
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   399
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   398
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   398
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   397
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   397
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   396
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   396
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   395
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   395
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   394
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   394
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   393
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   393
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   392
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   392
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   391
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   391
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   390
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   390
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   389
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   389
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   388
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   388
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   387
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   387
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   386
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   386
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   385
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   385
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   384
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   384
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   383
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   383
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   382
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   382
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   381
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   381
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   380
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   380
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   379
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   379
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   378
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   378
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   377
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   377
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   376
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   376
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   375
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   375
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   374
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   374
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   373
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   373
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   372
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   372
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   371
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   371
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   370
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   370
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   369
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   369
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   368
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   368
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   367
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   367
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   366
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   366
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   365
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   365
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   364
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   364
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   363
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   363
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   362
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   362
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   361
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   361
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   360
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   360
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   359
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   359
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   358
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   358
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   357
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   357
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   356
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   356
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   355
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   355
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   354
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   354
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   353
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   353
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   352
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   352
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   351
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   351
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   350
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   350
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   349
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   349
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   348
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   348
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   347
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   347
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   346
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   346
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   345
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   345
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   344
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   344
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   343
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   343
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   342
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   342
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   341
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   341
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   340
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   340
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   339
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   339
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   338
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   338
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   337
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   337
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   336
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   336
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   335
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   335
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   334
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   334
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   333
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   333
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   332
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   332
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   331
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   331
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   330
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   330
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   329
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   329
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   328
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   328
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   327
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   327
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   326
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   326
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   325
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   325
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   324
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   324
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   323
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   323
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   322
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   322
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   321
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   321
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   320
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   320
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   319
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   319
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   318
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   318
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   317
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   317
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   316
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   316
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   315
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   315
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   314
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   314
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   313
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   313
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   312
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   312
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   311
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   311
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   310
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   310
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   309
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   309
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   308
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   308
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   307
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   307
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   306
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   306
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   305
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   305
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   304
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   304
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   303
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   303
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   302
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   302
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   301
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   301
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   300
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   300
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   299
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   299
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   298
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   298
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   297
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   297
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   296
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   296
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   295
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   295
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   294
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   294
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   293
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   293
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   292
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   292
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   291
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   291
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   290
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   290
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   289
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   289
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   288
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   288
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   287
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   287
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   286
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   286
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   285
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   285
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   284
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   284
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   283
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   283
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   282
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   282
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   281
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   281
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   280
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   280
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   279
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   279
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   278
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   278
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   277
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   277
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   276
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   276
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   275
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   275
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   274
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   274
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   273
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   273
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   272
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   272
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   271
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   271
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   270
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   270
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   269
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   269
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   268
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   268
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   267
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   267
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   266
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   266
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   265
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   265
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   264
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   264
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   263
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   263
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   262
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   262
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   261
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   261
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   260
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   260
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   259
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   259
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   258
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   258
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   257
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   257
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   256
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   256
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   255
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   255
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   254
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   254
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   253
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   253
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   252
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   252
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   251
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   251
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   250
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   250
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   249
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   249
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   248
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   248
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   247
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   247
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   246
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   246
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   245
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   245
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   244
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   244
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   243
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   243
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   242
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   242
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   241
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   241
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   240
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   240
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   239
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   239
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   238
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   238
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   237
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   237
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   236
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   236
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   235
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   235
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   234
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   234
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   233
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   233
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   232
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   232
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   231
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   231
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   230
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   230
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   229
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   229
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   228
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   228
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   227
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   227
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   226
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   226
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   225
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   225
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   224
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   224
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   223
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   223
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   222
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   222
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   221
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   221
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   220
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   220
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   219
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   219
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   218
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   218
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   217
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   217
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   216
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   216
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   215
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   215
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   214
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   214
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   213
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   213
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   212
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   212
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   211
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   211
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   210
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   210
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   209
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   209
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   208
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   208
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   207
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   207
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   206
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   206
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   205
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   205
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   204
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   204
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   203
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   203
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   202
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   202
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   201
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   201
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   200
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   200
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   199
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   199
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   198
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   198
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   197
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   197
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   196
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   196
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   195
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   195
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   194
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   194
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   193
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   193
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   192
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   192
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   191
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   191
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   190
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   190
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   189
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   189
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   188
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   188
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   187
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   187
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   186
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   186
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   185
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   185
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   184
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   184
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   183
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   183
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   182
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   182
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   181
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   181
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   180
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   180
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   179
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   179
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   178
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   178
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   177
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   177
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   176
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   176
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   175
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   175
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   174
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   174
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   173
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   173
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   172
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   172
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   171
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   171
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   170
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   170
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   169
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   169
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   168
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   168
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   167
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   167
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   166
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   166
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   165
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   165
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   164
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   164
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   163
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   163
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   162
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   162
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   161
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   161
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   160
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   160
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   159
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   159
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   158
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   158
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   157
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   157
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   156
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   156
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   155
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   155
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   154
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   154
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   153
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   153
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   152
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   152
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   151
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   151
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   150
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   150
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   149
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   149
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   148
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   148
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   147
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   147
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   146
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   146
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   145
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   145
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   144
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   144
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   143
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   143
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   142
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   142
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   141
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   141
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   140
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   140
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   139
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   139
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   138
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   138
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   137
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   137
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   136
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   136
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   135
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   135
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   134
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   134
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   133
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   133
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   132
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   132
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   131
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   131
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   130
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   130
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   129
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   129
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   128
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   128
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   127
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   127
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   126
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   126
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   125
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   125
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   124
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   124
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   123
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   123
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   122
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   122
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   121
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   121
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   120
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   120
      Top             =   2280
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   119
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   119
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   118
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   118
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   117
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   117
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   116
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   116
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   115
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   115
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   114
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   114
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   113
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   113
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   112
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   112
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   111
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   111
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   110
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   110
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   109
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   109
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   108
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   108
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   107
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   107
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   106
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   106
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   105
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   105
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   104
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   104
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   103
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   103
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   102
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   102
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   101
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   101
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   100
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   100
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   99
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   99
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   98
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   98
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   97
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   97
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   96
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   96
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   95
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   95
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   94
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   94
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   93
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   93
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   92
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   92
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   91
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   91
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   90
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   90
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   89
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   89
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   88
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   88
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   87
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   87
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   86
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   86
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   85
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   85
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   84
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   84
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   83
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   83
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   82
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   82
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   81
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   81
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   80
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   80
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   79
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   79
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   78
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   78
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   77
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   77
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   76
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   76
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   75
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   75
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   74
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   74
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   73
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   73
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   72
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   72
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   71
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   71
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   70
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   70
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   69
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   69
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   68
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   68
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   67
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   67
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   66
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   66
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   65
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   65
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   64
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   64
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   63
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   63
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   62
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   62
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   61
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   61
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   60
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   60
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   59
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   59
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   58
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   58
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   57
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   57
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   56
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   56
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   55
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   55
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   54
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   54
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   53
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   53
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   52
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   52
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   51
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   51
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   50
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   50
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   49
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   49
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   48
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   48
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   47
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   47
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   46
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   46
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   45
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   45
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   44
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   44
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   43
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   43
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   42
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   42
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   41
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   41
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   40
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   40
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   39
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   39
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   38
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   38
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   37
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   37
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   36
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   36
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   35
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   35
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   34
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   34
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   33
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   33
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   32
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   32
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   31
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   31
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   30
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   30
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   29
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   28
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   28
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   27
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   27
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   26
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   25
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   25
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   24
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   24
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   23
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   22
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   21
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   20
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   19
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   18
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   17
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   16
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Start Facing:"
      Height          =   1335
      Left            =   7800
      TabIndex        =   407
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Right Click:"
      Height          =   255
      Left            =   7560
      TabIndex        =   418
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click:"
      Height          =   255
      Left            =   7560
      TabIndex        =   417
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Left Click:"
      Height          =   255
      Left            =   7560
      TabIndex        =   416
      Top             =   5520
      Width           =   735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Walk Through Wall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   7680
      TabIndex        =   415
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Spot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7680
      TabIndex        =   412
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "End Of maze"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7680
      TabIndex        =   414
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7680
      TabIndex        =   413
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Wall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   411
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Double Click To Open:"
      Height          =   255
      Left            =   7560
      TabIndex        =   409
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   255
      Left            =   7560
      TabIndex        =   402
      Top             =   3120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ST As Boolean
Dim D As Integer
Private Sub Command1_Click()
Dim Maze As TheMap
b = 1
'turn the maze into numbers so that it can be put into the file
For i = 0 To 399
If Picture1(i).BackColor = vbWhite Then a = "1"
If Picture1(i).BackColor = vbBlue Then
    'save starting position
    a = "1"
    start = i
End If
If Picture1(i).BackColor = vbBlack Then a = "0" 'wall
If Picture1(i).BackColor = vbCyan Then a = "3" ' secret
If Picture1(i).BackColor = vbRed Then a = "2": Tend = True ' finish
temp = temp + a
If Len(temp) = 20 Then Maze.Map(b) = temp: b = b + 1: temp = ""
Next i
'checks to see if the user remebered to put a start and finish in
If start = 0 Then g = MsgBox("Need A Starting Position!", 48, "Error"): Exit Sub
If Tend = False Then g = MsgBox("Maze Has No End!", 48, "Error"): Exit Sub
' put the starting position in x(w) and y(h)
h = Int((start / 20) + 1)
w = (start Mod 20) + 1
If h < 10 Then h = "0" & h
If w < 10 Then w = "0" & w
'save the starting position and pasword if needed
Maze.Map(0) = h & w & D & Trim(Text1.Text) & "|" & Space(10 - Len(Trim(Text1.Text))) & "0000"
Open App.Path & "\" & Trim(Text2.Text) & ".map" For Random As #1 Len = Len(Maze)
'save it all to the file
Put #1, 1, Maze
Close #1
File.Refresh
End Sub




Private Sub Command2_Click()

End Sub

Private Sub File_DblClick()
'opens the file if
Dim Maze As TheMap
Open App.Path & "\" & File.List(File.ListIndex) For Random As #1 Len = Len(Maze)
Get #1, 1, Maze
Close #1
h = Mid(Maze.Map(0), 1, 2)
w = Mid(Maze.Map(0), 3, 2)
D = Mid(Maze.Map(0), 5, 1)
Option1(D).Value = True
ST = True
Pass = Mid(Maze.Map(0), 6)
Text1.Text = ""
If Mid(Pass, 1, 1) <> "|" Then
'if there is a password, ask for it
    Pass = Left(Pass, InStr(1, Pass, "|") - 1)
    TPass = InputBox("Enter The Password:", "Password ")
    If TPass <> Pass Then g = MsgBox("Incorrect Password, Access Denied!", 48, "Error"): Exit Sub
    ' if the user doesn't know it, don't open the file
    Text1.Text = Trim(Pass)
End If
Text2.Text = Left(File.List(File.ListIndex), Len(File.List(File.ListIndex)) - 4)
For i = 1 To 20
For b = 1 To 20
    'load file
    a = Mid(Maze.Map(i), b, 1)
    If a = "0" Then Picture1(c).BackColor = vbBlack 'wall
    If a = "1" Then Picture1(c).BackColor = vbWhite 'open
    If a = "2" Then Picture1(c).BackColor = vbRed 'finish
    If a = "3" Then Picture1(c).BackColor = vbCyan 'secret
    c = c + 1
Next b
Next i
'put the start in
Picture1(Val(h) * 20 - 21 + Val(w)).BackColor = vbBlue
End Sub


Private Sub File_KeyDown(KeyCode As Integer, Shift As Integer)
'delete file if the delete button is pressed
If KeyCode = 46 Then Kill App.Path & "\" & File.List(File.ListIndex)
File.Refresh
End Sub


Private Sub Form_Load()
File.Path = App.Path
File.Refresh
End Sub

Private Sub Option1_Click(Index As Integer)
' save the direction the player will start
D = Index
End Sub

Private Sub Picture1_DblClick(Index As Integer)
'if double clicked then either make it red(finish) or blue(start)
If ST = True Then Picture1(Index).BackColor = vbRed
If ST = False Then Picture1(Index).BackColor = vbBlue: ST = True

End Sub


Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Picture1(Index).BackColor = vbBlue Then ST = False
'changes the space type
If Picture1(Index).BackColor = vbWhite And Button <> 2 Then
    Picture1(Index).BackColor = vbBlack
ElseIf Button <> 2 Then
    Picture1(Index).BackColor = vbWhite
ElseIf Button = 2 Then
    Picture1(Index).BackColor = vbCyan
End If

End Sub

Private Sub Text1_Change()
'check to see if password is to long
If Len(Text1.Text) > 10 Then
    Text1.Text = Left(Text1.Text, 10)
    MsgBox "Password To Long!"
End If
End Sub


