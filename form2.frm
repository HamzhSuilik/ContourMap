VERSION 5.00
Begin VB.Form form2 
   Caption         =   "Form2"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19035
   LinkTopic       =   "Form2"
   ScaleHeight     =   9645
   ScaleWidth      =   19035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   15840
      TabIndex        =   135
      Text            =   "Text3"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   15360
      Top             =   4800
   End
   Begin VB.TextBox dir7 
      Height          =   285
      Left            =   960
      TabIndex        =   134
      Text            =   "0"
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox dir8 
      Height          =   285
      Left            =   960
      TabIndex        =   133
      Text            =   "0"
      Top             =   8280
      Width           =   615
   End
   Begin VB.TextBox dir3 
      Height          =   285
      Left            =   960
      TabIndex        =   132
      Text            =   "0"
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox dir6 
      Height          =   285
      Left            =   960
      TabIndex        =   131
      Text            =   "0"
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox dir2 
      Height          =   285
      Left            =   960
      TabIndex        =   130
      Text            =   "0"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox dir4 
      Height          =   285
      Left            =   960
      TabIndex        =   129
      Text            =   "0"
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox dir5 
      Height          =   285
      Left            =   960
      TabIndex        =   128
      Text            =   "0"
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox dir1 
      Height          =   285
      Left            =   960
      TabIndex        =   127
      Text            =   "0"
      Top             =   6600
      Width           =   615
   End
   Begin VB.Timer Timer6 
      Interval        =   120
      Left            =   2520
      Top             =   6600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   126
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   125
      Text            =   "Text2"
      Top             =   8640
      Width           =   855
   End
   Begin VB.TextBox virend 
      Height          =   285
      Left            =   4680
      TabIndex        =   124
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox horend 
      Height          =   285
      Left            =   3720
      TabIndex        =   123
      Top             =   7440
      Width           =   735
   End
   Begin VB.TextBox mvir 
      Height          =   285
      Left            =   4680
      TabIndex        =   122
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox mhor 
      Height          =   375
      Left            =   3720
      TabIndex        =   121
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox hor 
      Height          =   375
      Left            =   3720
      TabIndex        =   120
      Text            =   "0"
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox vir 
      Height          =   375
      Left            =   4680
      TabIndex        =   119
      Text            =   "0"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      TabIndex        =   118
      Text            =   "Text1"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   840
      Top             =   5400
   End
   Begin VB.TextBox bb4 
      Height          =   285
      Left            =   1560
      TabIndex        =   117
      Text            =   "252"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox bbb4 
      Height          =   285
      Left            =   2400
      TabIndex        =   116
      Text            =   "300"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox num4 
      Height          =   285
      Left            =   3120
      TabIndex        =   115
      Text            =   "000"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox numm4 
      Height          =   285
      Left            =   4320
      TabIndex        =   114
      Text            =   "0"
      Top             =   5520
      Width           =   375
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1080
      Top             =   4680
   End
   Begin VB.TextBox bb3 
      Height          =   285
      Left            =   1800
      TabIndex        =   113
      Text            =   "51"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox bbb3 
      Height          =   285
      Left            =   2520
      TabIndex        =   112
      Text            =   "210"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox num3 
      Height          =   285
      Left            =   3360
      TabIndex        =   111
      Text            =   "000"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox numm3 
      Height          =   285
      Left            =   4560
      TabIndex        =   110
      Text            =   "0"
      Top             =   4800
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   960
      Top             =   3840
   End
   Begin VB.TextBox bb2 
      Height          =   285
      Left            =   1680
      TabIndex        =   109
      Text            =   "250"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox bbb2 
      Height          =   285
      Left            =   2400
      TabIndex        =   108
      Text            =   "251"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox num2 
      Height          =   285
      Left            =   3240
      TabIndex        =   107
      Text            =   "000"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox numm2 
      Height          =   285
      Left            =   4440
      TabIndex        =   106
      Text            =   "0"
      Top             =   3960
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   1455
      Left            =   9600
      Picture         =   "form2.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   105
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox b9 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   104
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox b10 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   103
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox c2 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   102
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox pic9_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   101
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox b7 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   100
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox b8 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   99
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic8_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   98
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic7_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   97
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox b5 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   96
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox b6 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   95
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic6_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   94
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic5_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   93
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox b3 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   92
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox b4 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   91
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic4_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   90
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic3_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   89
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox a10 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   88
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox b2 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   87
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox pic2_9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   86
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox a9 
      Height          =   495
      Left            =   11760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   85
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox pic9_8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   84
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox c3 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   83
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox c4 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   82
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox pic9_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   81
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox pic9_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   80
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox c5 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   79
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox c6 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   78
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox pic9_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   77
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox pic9_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   76
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox c7 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   75
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox c8 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   74
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox pic9_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   73
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox pic9_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   72
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox c9 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   71
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox c10 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   70
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox d2 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   69
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox pic7_8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   68
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic8_8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   67
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic8_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   66
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic7_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   65
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic7_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   64
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic8_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   63
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic8_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   62
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic7_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   61
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic5_8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   60
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic6_8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   59
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic6_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   58
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic5_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   57
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic5_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   56
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic6_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   55
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic6_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   54
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic5_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   53
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic7_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   52
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic8_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   51
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic8_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   50
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox pic7_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   49
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic7_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   48
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic8_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   47
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox d3 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   46
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox d4 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox pic5_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   44
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic6_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic6_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   42
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pic5_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic5_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   40
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic6_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox d5 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox d6 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   37
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox pic3_8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   36
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox pic4_28 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   35
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic4_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   34
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic3_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   33
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox pic3_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   32
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox pic4_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic4_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic3_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox a8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   28
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox pic2_8 
      Height          =   495
      Left            =   11160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox pic2_7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   26
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox a7 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   25
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox a6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   24
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox pic2_6 
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox pic2_5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox a5 
      Height          =   495
      Left            =   9360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox pic3_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox pic4_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic4_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox pic3_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox pic3_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox pic4_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox d7 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox d8 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox a4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox pic2_4 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox pic2_3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox a3 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox numm1 
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Text            =   "0"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox num1 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "000"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox bbb1 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "250"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox bb1 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "50"
      Top             =   3240
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   3120
   End
   Begin VB.TextBox elevation 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Text            =   "201"
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox a1 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox d9 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox pic2_2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox a2 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   2760
      Width           =   495
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text8_Change()

End Sub



Private Sub Timer1_Timer()
Dim agree As Integer
Dim min As Double
Dim max As Double
Dim location As Double
Dim div As Double

min = 0
max = 0
agree = 1
location = 0

If elevation.Text > bb1.Text And elevation.Text < bbb1.Text Then agree = 1
If elevation.Text < bb1.Text And elevation.Text > bbb1.Text Then agree = 1

Text1.Text = agree

If agree = 1 Then
If bb1 > bbb1 Then
max = bb1.Text
min = bbb1.Text
Else
max = bbb1.Text
min = bb1.Text
End If

div = (max - min)
location = 9 * (elevation.Text - min)

num1.Text = location / div
 
If min = bb1 Then
If num1.Text > 0 And num1.Text < 0.5 Then numm1.Text = 1
If num1.Text > 0.5 And num1.Text < 1.5 Then numm1.Text = 2
If num1.Text > 1.5 And num1.Text < 2.5 Then numm1.Text = 3
If num1.Text > 2.5 And num1.Text < 3.5 Then numm1.Text = 4
If num1.Text > 3.5 And num1.Text < 4.5 Then numm1.Text = 5
If num1.Text > 4.5 And num1.Text < 5.5 Then numm1.Text = 6
If num1.Text > 5.5 And num1.Text < 6.5 Then numm1.Text = 7
If num1.Text > 6.5 And num1.Text < 7.5 Then numm1.Text = 8
If num1.Text > 7.5 And num1.Text < 8.5 Then numm1.Text = 9
If num1.Text > 8.5 And num1.Text < 9 Then numm1.Text = 10
End If

If min = bbb1 Then
If num1.Text > 0 And num1.Text < 0.5 Then numm1.Text = 10
If num1.Text > 0.5 And num1.Text < 1.5 Then numm1.Text = 9
If num1.Text > 1.5 And num1.Text < 2.5 Then numm1.Text = 8
If num1.Text > 2.5 And num1.Text < 3.5 Then numm1.Text = 7
If num1.Text > 3.5 And num1.Text < 4.5 Then numm1.Text = 6
If num1.Text > 4.5 And num1.Text < 5.5 Then numm1.Text = 5
If num1.Text > 5.5 And num1.Text < 6.5 Then numm1.Text = 4
If num1.Text > 6.5 And num1.Text < 7.5 Then numm1.Text = 3
If num1.Text > 7.5 And num1.Text < 8.5 Then numm1.Text = 2
If num1.Text > 8.5 And num1.Text < 9 Then numm1.Text = 1
Text1.Text = 74387

End If

If numm1.Text = 1 Then a1.Picture = Pic.Picture
If numm1.Text = 2 Then a2.Picture = Pic.Picture
If numm1.Text = 3 Then a3.Picture = Pic.Picture
If numm1.Text = 4 Then a4.Picture = Pic.Picture
If numm1.Text = 5 Then a5.Picture = Pic.Picture
If numm1.Text = 6 Then a6.Picture = Pic.Picture
If numm1.Text = 7 Then a7.Picture = Pic.Picture
If numm1.Text = 8 Then a8.Picture = Pic.Picture
If numm1.Text = 9 Then a9.Picture = Pic.Picture
If numm1.Text = 10 Then a10.Picture = Pic.Picture


End If
 
 
 
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Dim agree As Integer
Dim min As Double
Dim max As Double
Dim location As Double
Dim div As Double

min = 0
max = 0
agree = 1
location = 0

If elevation.Text > bb2.Text And elevation.Text < bbb2.Text Then agree = 1
If elevation.Text < bb2.Text And elevation.Text > bbb2.Text Then agree = 1

If agree = 1 Then
If bb2 > bbb2 Then
max = bb2.Text
min = bbb2.Text
Else
max = bbb2.Text
min = bb2.Text
End If

div = (max - min)
location = 9 * (elevation.Text - min)

num2.Text = location / div
 
If min = bb2 Then
If num2.Text > 0 And num2.Text < 0.5 Then numm2.Text = 1
If num2.Text > 0.5 And num2.Text < 1.5 Then numm2.Text = 2
If num2.Text > 1.5 And num2.Text < 2.5 Then numm2.Text = 3
If num2.Text > 2.5 And num2.Text < 3.5 Then numm2.Text = 4
If num2.Text > 3.5 And num2.Text < 4.5 Then numm2.Text = 5
If num2.Text > 4.5 And num2.Text < 5.5 Then numm2.Text = 6
If num2.Text > 5.5 And num2.Text < 6.5 Then numm2.Text = 7
If num2.Text > 6.5 And num2.Text < 7.5 Then numm2.Text = 8
If num2.Text > 7.5 And num2.Text < 8.5 Then numm2.Text = 9
If num2.Text > 8.5 And num2.Text < 9 Then numm2.Text = 10

End If

If min = bbb2 Then
If num2.Text > 0 And num2.Text < 0.5 Then numm2.Text = 10
If num2.Text > 0.5 And num2.Text < 1.5 Then numm2.Text = 9
If num2.Text > 1.5 And num2.Text < 2.5 Then numm2.Text = 8
If num2.Text > 2.5 And num2.Text < 3.5 Then numm2.Text = 7
If num2.Text > 3.5 And num2.Text < 4.5 Then numm2.Text = 6
If num2.Text > 4.5 And num2.Text < 5.5 Then numm2.Text = 5
If num2.Text > 5.5 And num2.Text < 6.5 Then numm2.Text = 4
If num2.Text > 6.5 And num2.Text < 7.5 Then numm2.Text = 3
If num2.Text > 7.5 And num2.Text < 8.5 Then numm2.Text = 2
If num2.Text > 8.5 And num2.Text < 9 Then numm2.Text = 1

End If

If numm2.Text = 1 Then a10.Picture = Pic.Picture
If numm2.Text = 2 Then b2.Picture = Pic.Picture
If numm2.Text = 3 Then b3.Picture = Pic.Picture
If numm2.Text = 4 Then b4.Picture = Pic.Picture
If numm2.Text = 5 Then b5.Picture = Pic.Picture
If numm2.Text = 6 Then b6.Picture = Pic.Picture
If numm2.Text = 7 Then b7.Picture = Pic.Picture
If numm2.Text = 8 Then b8.Picture = Pic.Picture
If numm2.Text = 9 Then b9.Picture = Pic.Picture
If numm2.Text = 10 Then b10.Picture = Pic.Picture


End If
 
 
 
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
Dim agree As Integer
Dim min As Double
Dim max As Double
Dim location As Double
Dim div As Double

min = 0
max = 0
agree = 1
location = 0

If elevation.Text > bb3.Text And elevation.Text < bbb3.Text Then agree = 1
If elevation.Text < bb3.Text And elevation.Text > bbb3.Text Then agree = 1

If agree = 1 Then
If bb3 > bbb3 Then
max = bb3.Text
min = bbb3.Text
Else
max = bbb3.Text
min = bb3.Text
End If

div = (max - min)
location = 9 * (elevation.Text - min)

num3.Text = location / div
 
If min = bb3 Then
If num3.Text > 0 And num3.Text < 0.5 Then numm3.Text = 1
If num3.Text > 0.5 And num3.Text < 1.5 Then numm3.Text = 2
If num3.Text > 1.5 And num3.Text < 2.5 Then numm3.Text = 3
If num3.Text > 2.5 And num3.Text < 3.5 Then numm3.Text = 4
If num3.Text > 3.5 And num3.Text < 4.5 Then numm3.Text = 5
If num3.Text > 4.5 And num3.Text < 5.5 Then numm3.Text = 6
If num3.Text > 5.5 And num3.Text < 6.5 Then numm3.Text = 7
If num3.Text > 6.5 And num3.Text < 7.5 Then numm3.Text = 8
If num3.Text > 7.5 And num3.Text < 8.5 Then numm3.Text = 9
If num3.Text > 8.5 And num3.Text < 9 Then numm3.Text = 10

End If

If min = bbb3 Then
If num3.Text > 0 And num3.Text < 0.5 Then numm3.Text = 10
If num3.Text > 0.5 And num3.Text < 1.5 Then numm3.Text = 9
If num3.Text > 1.5 And num3.Text < 2.5 Then numm3.Text = 8
If num3.Text > 2.5 And num3.Text < 3.5 Then numm3.Text = 7
If num3.Text > 3.5 And num3.Text < 4.5 Then numm3.Text = 6
If num3.Text > 4.5 And num3.Text < 5.5 Then numm3.Text = 5
If num3.Text > 5.5 And num3.Text < 6.5 Then numm3.Text = 4
If num3.Text > 6.5 And num3.Text < 7.5 Then numm3.Text = 3
If num3.Text > 7.5 And num3.Text < 8.5 Then numm3.Text = 2
If num3.Text > 8.5 And num3.Text < 9 Then numm3.Text = 1

End If

If numm3.Text = 1 Then b10.Picture = Pic.Picture
If numm3.Text = 2 Then c2.Picture = Pic.Picture
If numm3.Text = 3 Then c3.Picture = Pic.Picture
If numm3.Text = 4 Then c4.Picture = Pic.Picture
If numm3.Text = 5 Then c5.Picture = Pic.Picture
If numm3.Text = 6 Then c6.Picture = Pic.Picture
If numm3.Text = 7 Then c7.Picture = Pic.Picture
If numm3.Text = 8 Then c8.Picture = Pic.Picture
If numm3.Text = 9 Then c9.Picture = Pic.Picture
If numm3.Text = 10 Then c10.Picture = Pic.Picture


End If
 
 
 
Timer3.Enabled = False

End Sub

Private Sub Timer4_Timer()
Dim agree As Integer
Dim min As Double
Dim max As Double
Dim location As Double
Dim div As Double

min = 0
max = 0
agree = 1
location = 0

If elevation.Text > bb4.Text And elevation.Text < bbb4.Text Then agree = 1
If elevation.Text < bb4.Text And elevation.Text > bbb4.Text Then agree = 1

If agree = 1 Then
If bb4 > bbb4 Then
max = bb4.Text
min = bbb4.Text
Else
max = bbb4.Text
min = bb4.Text
End If

div = (max - min)
location = 9 * (elevation.Text - min)

num4.Text = location / div
 
If min = bb4 Then
If num4.Text > 0 And num4.Text < 0.5 Then numm4.Text = 1
If num4.Text > 0.5 And num4.Text < 1.5 Then numm4.Text = 2
If num4.Text > 1.5 And num4.Text < 2.5 Then numm4.Text = 3
If num4.Text > 2.5 And num4.Text < 3.5 Then numm4.Text = 4
If num4.Text > 3.5 And num4.Text < 4.5 Then numm4.Text = 5
If num4.Text > 4.5 And num4.Text < 5.5 Then numm4.Text = 6
If num4.Text > 5.5 And num4.Text < 6.5 Then numm4.Text = 7
If num4.Text > 6.5 And num4.Text < 7.5 Then numm4.Text = 8
If num4.Text > 7.5 And num4.Text < 8.5 Then numm4.Text = 9
If num4.Text > 8.5 And num4.Text < 9 Then numm4.Text = 10

End If

If min = bbb4 Then
If num4.Text > 0 And num4.Text < 0.5 Then numm4.Text = 10
If num4.Text > 0.5 And num4.Text < 1.5 Then numm4.Text = 9
If num4.Text > 1.5 And num4.Text < 2.5 Then numm4.Text = 8
If num4.Text > 2.5 And num4.Text < 3.5 Then numm4.Text = 7
If num4.Text > 3.5 And num4.Text < 4.5 Then numm4.Text = 6
If num4.Text > 4.5 And num4.Text < 5.5 Then numm4.Text = 5
If num4.Text > 5.5 And num4.Text < 6.5 Then numm4.Text = 4
If num4.Text > 6.5 And num4.Text < 7.5 Then numm4.Text = 3
If num4.Text > 7.5 And num4.Text < 8.5 Then numm4.Text = 2
If num4.Text > 8.5 And num4.Text < 9 Then numm4.Text = 1

End If

If numm4.Text = 1 Then c10.Picture = Pic.Picture
If numm4.Text = 2 Then d2.Picture = Pic.Picture
If numm4.Text = 3 Then d3.Picture = Pic.Picture
If numm4.Text = 4 Then d4.Picture = Pic.Picture
If numm4.Text = 5 Then d5.Picture = Pic.Picture
If numm4.Text = 6 Then d6.Picture = Pic.Picture
If numm4.Text = 7 Then d7.Picture = Pic.Picture
If numm4.Text = 8 Then d8.Picture = Pic.Picture
If numm4.Text = 9 Then d9.Picture = Pic.Picture
If numm4.Text = 10 Then a1.Picture = Pic.Picture


End If
 
 
 Timer4.Enabled = False


End Sub



Private Sub Timer5_Timer()


Dim mohor As Integer
Dim movir As Integer
Dim step As Integer
Dim start As Integer
Dim control As Integer

start = 1

If (mhor.Text ^ 2) ^ 0.5 > (mvir.Text ^ 2) ^ 0.5 Then control = 1
If (mhor.Text ^ 2) ^ 0.5 < (mvir.Text ^ 2) ^ 0.5 Then control = 2

If (mhor.Text ^ 2) ^ 0.5 > (mvir.Text ^ 2) ^ 0.5 Then step = (mhor.Text ^ 2) ^ 0.5
If (mhor.Text ^ 2) ^ 0.5 < (mvir.Text ^ 2) ^ 0.5 Then step = (mvir.Text ^ 2) ^ 0.5


If mhor > 0 Then
mohor = 1
Else
mohor = -1
End If
If mhor = 0 Then mohor = 0

If mvir > 0 Then
movir = 1
Else
movir = -1
End If
If mvir = 0 Then movir = 0


Do While step <> (start - 1)

If start <> dir1.Text And start <> dir2.Text And start <> dir3.Text And start <> dir4.Text And start <> dir5.Text And start <> dir6.Text And start <> dir7.Text And start <> dir8.Text Then
 hor.Text = hor.Text + mohor
 vir.Text = vir.Text + movir
 
 Else
 If control = 1 Then hor.Text = hor.Text + mohor
 If control = 2 Then vir.Text = vir.Text + movir

End If


If hor.Text = 2 And vir.Text = 2 Then pic2_2.Picture = Pic.Picture
If hor.Text = 2 And vir.Text = 3 Then pic2_3.Picture = Pic.Picture
If hor.Text = 2 And vir.Text = 4 Then pic2_4.Picture = Pic.Picture
If hor.Text = 2 And vir.Text = 5 Then pic2_5.Picture = Pic.Picture
If hor.Text = 2 And vir.Text = 6 Then pic2_6.Picture = Pic.Picture
If hor.Text = 2 And vir.Text = 7 Then pic2_7.Picture = Pic.Picture
If hor.Text = 2 And vir.Text = 8 Then pic2_8.Picture = Pic.Picture
If hor.Text = 2 And vir.Text = 9 Then pic2_9.Picture = Pic.Picture



If hor.Text = 3 And vir.Text = 2 Then pic3_2.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 3 Then pic3_3.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 4 Then pic3_4.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 5 Then pic3_5.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 6 Then pic3_6.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 7 Then pic3_7.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 8 Then pic3_8.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 9 Then pic3_9.Picture = Pic.Picture




If hor.Text = 4 And vir.Text = 2 Then pic4_2.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 3 Then pic4_3.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 4 Then pic4_4.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 5 Then pic4_5.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 6 Then pic4_6.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 7 Then pic4_7.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 8 Then pic4_8.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 9 Then pic4_9.Picture = Pic.Picture




If hor.Text = 5 And vir.Text = 2 Then pic5_2.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 3 Then pic5_3.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 4 Then pic5_4.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 5 Then pic5_5.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 6 Then pic5_6.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 7 Then pic5_7.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 8 Then pic5_8.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 9 Then pic5_9.Picture = Pic.Picture




If hor.Text = 6 And vir.Text = 2 Then pic6_2.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 3 Then pic6_3.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 4 Then pic6_4.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 5 Then pic6_5.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 6 Then pic6_6.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 7 Then pic6_7.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 8 Then pic6_8.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 9 Then pic6_9.Picture = Pic.Picture




If hor.Text = 7 And vir.Text = 2 Then pic7_2.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 3 Then pic7_3.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 4 Then pic7_4.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 5 Then pic7_5.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 6 Then pic7_6.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 7 Then pic7_7.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 8 Then pic7_8.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 9 Then pic7_9.Picture = Pic.Picture



If hor.Text = 8 And vir.Text = 2 Then pic8_2.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 3 Then pic8_3.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 4 Then pic8_4.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 5 Then pic8_5.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 6 Then pic8_6.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 7 Then pic8_7.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 8 Then pic8_8.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 9 Then pic8_9.Picture = Pic.Picture



If hor.Text = 9 And vir.Text = 2 Then pic9_2.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 3 Then pic9_3.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 4 Then pic9_4.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 5 Then pic9_5.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 6 Then pic9_6.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 7 Then pic9_7.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 8 Then pic9_8.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 9 Then pic9_9.Picture = Pic.Picture

If hor.Text = 1 And vir.Text = 1 Then a1.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 2 Then a2.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 3 Then a3.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 4 Then a4.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 5 Then a5.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 6 Then a6.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 7 Then a7.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 8 Then a8.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 9 Then a9.Picture = Pic.Picture
If hor.Text = 1 And vir.Text = 10 Then a10.Picture = Pic.Picture





If hor.Text = 2 And vir.Text = 10 Then b2.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 10 Then b3.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 10 Then b4.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 10 Then b5.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 10 Then b6.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 10 Then b7.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 10 Then b8.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 10 Then b8.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 10 Then b8.Picture = Pic.Picture


If hor.Text = 10 And vir.Text = 1 Then c10.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 2 Then c9.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 3 Then c8.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 4 Then c7.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 5 Then c6.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 6 Then c5.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 7 Then c4.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 8 Then c3.Picture = Pic.Picture
If hor.Text = 10 And vir.Text = 9 Then c2.Picture = Pic.Picture



If hor.Text = 2 And vir.Text = 1 Then d9.Picture = Pic.Picture
If hor.Text = 3 And vir.Text = 1 Then d8.Picture = Pic.Picture
If hor.Text = 4 And vir.Text = 1 Then d7.Picture = Pic.Picture
If hor.Text = 5 And vir.Text = 1 Then d6.Picture = Pic.Picture
If hor.Text = 6 And vir.Text = 1 Then d5.Picture = Pic.Picture
If hor.Text = 7 And vir.Text = 1 Then d4.Picture = Pic.Picture
If hor.Text = 8 And vir.Text = 1 Then d3.Picture = Pic.Picture
If hor.Text = 9 And vir.Text = 1 Then d2.Picture = Pic.Picture






start = start + 1
Loop


Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()

If numm1.Text <> 0 Then
hor.Text = 1
vir.Text = numm1.Text
End If

If numm2.Text <> 0 Then
If vir.Text = 0 Then
hor.Text = numm2.Text
vir.Text = 10
Else
horend.Text = numm2.Text
virend.Text = 10
End If
End If

If numm3.Text <> 0 Then
If vir.Text = 0 Then
hor.Text = 10
vir.Text = 11 - numm3.Text
Else
horend.Text = 10
virend.Text = 11 - numm3.Text
End If
End If

If numm4.Text <> 0 Then
horend.Text = 11 - numm4.Text
virend.Text = 1
End If

mhor.Text = horend.Text - (hor.Text)
mvir.Text = virend.Text - (vir.Text)

If mvir.Text > 0 Then mvir.Text = mvir.Text - 1
If mvir.Text < 0 Then mvir.Text = mvir.Text + 1

If mhor.Text > 0 Then mhor.Text = mhor.Text - 1
If mhor.Text < 0 Then mhor.Text = mhor.Text + 1



If (mhor.Text ^ 2) ^ 0.5 <> (mvir.Text ^ 2) ^ 0.5 Then





Text2.Text = (mhor.Text ^ 2) ^ 0.5 - (mvir.Text ^ 2) ^ 0.5

Dim a As Double

If (mhor.Text ^ 2) ^ 0.5 > (mvir.Text ^ 2) ^ 0.5 Then
a = (mhor.Text ^ 2) ^ 0.5
Else
a = (mvir.Text ^ 2) ^ 0.5
End If

Dim c As Integer
Dim b As Integer
b = a / ((Text2.Text ^ 2) ^ 0.5 + 1)



c = a / ((Text2.Text ^ 2) ^ 0.5 + 0)
Text3.Text = c
dir1.Text = c
If dir1.Text = 0 Then dir1.Text = 1



If ((Text2.Text ^ 2) ^ 0.5) > 1 Then
c = (a * 2) / ((Text2.Text ^ 2) ^ 0.5 + 0)
dir2.Text = c
Do While (dir2.Text = dir1.Text Or dir2.Text = 0)
dir2.Text = dir2.Text + 1
If dir2.Text = 9 Then dir2.Text = 1
Loop
End If




If ((Text2.Text ^ 2) ^ 0.5) > 2 Then
c = (a * 3) / ((Text2.Text ^ 2) ^ 0.5 + 0)
dir3.Text = c


Do While (dir3.Text = dir1.Text Or dir3.Text = 0 Or dir3.Text = dir2.Text)
dir3.Text = dir3.Text + 1
If dir3.Text = 9 Then dir3.Text = 1
Loop

End If


If ((Text2.Text ^ 2) ^ 0.5) > 3 Then
c = (a * 4) / ((Text2.Text ^ 2) ^ 0.5 + 0)
dir4.Text = c


Do While (dir4.Text = dir1.Text Or dir4.Text = 0 Or dir4.Text = dir2.Text Or dir4.Text = dir3.Text)
dir4.Text = dir4.Text + 1
If dir4.Text = 9 Then dir4.Text = 1
Loop

End If


If ((Text2.Text ^ 2) ^ 0.5) > 4 Then
c = (a * 5) / ((Text2.Text ^ 2) ^ 0.5 + 0)
dir5.Text = c


Do While (dir5.Text = dir1.Text Or dir5.Text = 0 Or dir5.Text = dir2.Text Or dir5.Text = dir3.Text Or dir5.Text = dir4.Text)
dir5.Text = dir5.Text + 1
If dir5.Text = 9 Then dir5.Text = 1
Loop

End If


If ((Text2.Text ^ 2) ^ 0.5) > 5 Then
c = (a * 6) / ((Text2.Text ^ 2) ^ 0.5 + 0)
dir6.Text = c


Do While (dir6.Text = dir1.Text Or dir6.Text = 0 Or dir6.Text = dir2.Text Or dir6.Text = dir3.Text Or dir6.Text = dir4.Text Or dir6.Text = dir5.Text)
dir6.Text = dir6.Text + 1
If dir6.Text = 9 Then dir6.Text = 1
Loop

End If



If ((Text2.Text ^ 2) ^ 0.5) > 6 Then
c = (a * 7) / ((Text2.Text ^ 2) ^ 0.5 + 0)
dir7.Text = c


Do While (dir7.Text = dir1.Text Or dir7.Text = 0 Or dir7.Text = dir2.Text Or dir7.Text = dir3.Text Or dir7.Text = dir4.Text Or dir7.Text = dir5.Text Or dir7.Text = dir6.Text)
dir7.Text = dir7.Text + 1
If dir7.Text = 9 Then dir7.Text = 1
Loop

End If


If ((Text2.Text ^ 2) ^ 0.5) > 7 Then

c = (a * 8) \ ((Text2.Text ^ 2) ^ 0.5 + 0)
dir8.Text = c


Do While (dir8.Text = dir1.Text Or dir8.Text = 0 Or dir8.Text = dir2.Text Or dir8.Text = dir3.Text Or dir8.Text = dir4.Text Or dir8.Text = dir5.Text Or dir8.Text = dir6.Text Or dir8.Text = dir7.Text)
dir8.Text = dir8.Text + 1
If dir8.Text = 9 Then dir8.Text = 1
Loop


End If



















End If
Timer5.Enabled = True

Timer6.Enabled = False
End Sub



