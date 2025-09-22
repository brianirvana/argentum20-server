Attribute VB_Name = "ModMap"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Const MAX_RANDOM_TELEPORT_IN_MAP = 20

Private MapSize As t_MapSize
Private MapDat  As t_MapDat

Private Type t_Position
    x As Integer
    y As Integer
End Type

'Item type
Private Type t_Item
    ObjIndex As Integer
    amount As Integer
End Type

Private Type t_WorldPos
    Map As Integer
    x As Byte
    y As Byte
End Type

Private Type t_Grh
    GrhIndex As Long
    FrameCounter As Single
    Speed As Single
    Started As Byte
    alpha_blend As Boolean
    angle As Single
End Type

Private Type t_GrhData
    sX As Integer
    sY As Integer
    filenum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Integer
    Speed As Integer
    mini_map_color As Long
End Type

Private Type t_DatosTrigger
    x As Integer
    y As Integer
    trigger As Integer
End Type

Private Type t_DatosLuces
    x As Integer
    y As Integer
    Color As Long
    Rango As Byte
End Type

Private Type t_DatosParticulas
    x As Integer
    y As Integer
    Particula As Long
End Type

Private Type t_DatosNPC
    x As Integer
    y As Integer
    NpcIndex As Integer
End Type

Private Type t_DatosObjs
    x As Integer
    y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type t_DatosTE
    x As Integer
    y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type t_MapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type t_MapDat
    map_name As String
    backup_mode As Byte
    restrict_mode As String
    music_numberHi As Long
    music_numberLow As Long
    Seguro As Byte
    zone As String
    terrain As String
    ambient As String
    base_light As Long
    letter_grh As Long
    level As Long
    extra2 As Long
    Salida As String
    lluvia As Byte
    Nieve As Byte
    niebla As Byte
End Type

Private Type t_MapHeader
    NumeroBloqueados As Long
    NumeroLayers(1 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type t_DatosBloqueados
    x As Integer
    y As Integer
    Lados As Byte
End Type

Private Type t_DatosGrh
    x As Integer
    y As Integer
    GrhIndex As Long
End Type

Public Function get_map_name(ByVal Map As Long) As String

    On Error GoTo get_map_name_Err

    get_map_name = MapInfo(Map).map_name

    Exit Function

get_map_name_Err:
    Call TraceError(Err.Number, Err.Description, "Acciones.get_map_name", Erl)
End Function

'---------------------------------------------------------------------------------------
' Procedure : esCiudad
' Last Author : [/About] Brian Sabatier (brian.sabatier87@gmail.com - https://github.com/brianirvana/brianirvana)
' Last Date : 21/9/2025
' Purpose   :
'---------------------------------------------------------------------------------------

Public Function esCiudad(ByVal Map As Integer) As Boolean

Dim i                           As Byte

10  On Error GoTo esCiudad_Error

20  For i = 0 To UBound(TotalMapasCiudades)
30      If TotalMapasCiudades(i) = Map Then
40          esCiudad = True
50          Exit Function
60      End If
70  Next i

80  On Error GoTo 0
90  Exit Function

esCiudad_Error:
100 esCiudad = False
110 Call Logging.TraceError(Err.Number, Err.Description, "ModMap.esCiudad", Erl())

End Function

Public Function CanAddTrapAt(ByVal mapIndex As Integer, ByVal posX As Integer, ByVal posY As Integer) As Boolean

    If Not MapData(mapIndex, posX, posY).Trap Is Nothing Then
        Exit Function
    End If
    
    If MapData(mapIndex, posX, posY).Blocked Then Exit Function
    If MapData(mapIndex, posX, posY).npcIndex > 0 Then Exit Function
    If MapData(mapIndex, posX, posY).UserIndex > 0 Then Exit Function
    If MapData(mapIndex, posX, posY).ObjInfo.objIndex > 0 Then Exit Function
    CanAddTrapAt = True
    
End Function

Public Sub ActivateTrap(ByVal TargetIndex, ByVal TargetType As e_ReferenceType, ByVal map As Integer, ByVal posX As Integer, ByVal posY As Integer)

    On Error GoTo ActivateTrap_Error

10  If MapData(Map, PosX, PosY).Trap Is Nothing Then
20      Exit Sub
30  End If
40  If Not MapData(Map, PosX, PosY).Trap.CanAffectTarget(TargetIndex, TargetType) Then
50      Exit Sub
60  End If
70  Call MapData(Map, PosX, PosY).Trap.trigger(TargetIndex, TargetType)

    On Error GoTo 0
    Exit Sub

ActivateTrap_Error:

    Call Logging.TraceError(Err.Number, Err.Description, "ModMap.ActivateTrap", Erl())
End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal b As Byte)
    'b ahora es boolean,
    'b=true bloquea el tile en (x,y)
    'b=false desbloquea el tile en (x,y)
    'toMap = true -> Envia los datos a todo el mapa
    'toMap = false -> Envia los datos al user
    'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
    'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
    '  Uso bloqueo parcial
10  On Error GoTo Bloquear_Err
    ' Envío sólo los flags de bloq
20  b = b And e_Block.ALL_SIDES

30  If toMap Then
40      Call SendData(SendTarget.toMap, sndIndex, PrepareMessage_BlockPosition(x, y, b))
50  Else
60      Call Write_BlockPosition(sndIndex, x, y, b)
70  End If

80  Exit Sub

Bloquear_Err:
90  Call TraceError(Err.Number, Err.Description, "General.Bloquear", Erl)
End Sub

Sub BlockAndInform(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal NewState As Integer)

    On Error GoTo BlockAndInform_Error

10  If NewState Then
20      MapData(Map, x, y).Blocked = e_Block.ALL_SIDES Or e_Block.GM
30  Else
40      MapData(Map, x, y).Blocked = 0
50  End If

60  Call Bloquear(True, Map, x, y, MapData(Map, x, y).Blocked)

    On Error GoTo 0
    Exit Sub

BlockAndInform_Error:

    Call Logging.TraceError(Err.Number, Err.Description, "ModMap.BlockAndInform", Erl())

End Sub

Sub MostrarBloqueosPuerta(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal x As Integer, ByVal y As Integer)

Dim Map                         As Integer
Dim ModPuerta                   As Integer

10  On Error GoTo MostrarBloqueosPuerta_Err

20  If toMap Then
30      Map = sndIndex
40  Else
50      Map = UserList(sndIndex).pos.Map
60  End If

70  ModPuerta = ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Subtipo

80  Select Case ModPuerta
        Case 0
            ' Bloqueos superiores
90          Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
100         Call Bloquear(toMap, sndIndex, x - 1, y, MapData(Map, x - 1, y).Blocked)

            ' Bloqueos inferiores
110         Call Bloquear(toMap, sndIndex, x, y + 1, MapData(Map, x, y + 1).Blocked)
120         Call Bloquear(toMap, sndIndex, x - 1, y + 1, MapData(Map, x - 1, y + 1).Blocked)

130     Case 1
            ' para palancas o teclas sin modicar bloqueos en X,Y

140     Case 2
            ' Bloqueos superiores
150         Call Bloquear(toMap, sndIndex, x, y - 1, MapData(Map, x, y - 1).Blocked)
160         Call Bloquear(toMap, sndIndex, x - 1, y - 1, MapData(Map, x - 1, y - 1).Blocked)
170         Call Bloquear(toMap, sndIndex, x + 1, y - 1, MapData(Map, x + 1, y - 1).Blocked)
            ' Bloqueos inferiores
180         Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
190         Call Bloquear(toMap, sndIndex, x - 1, y, MapData(Map, x - 1, y).Blocked)
200         Call Bloquear(toMap, sndIndex, x + 1, y, MapData(Map, x + 1, y).Blocked)

210     Case 3
            ' Bloqueos superiores
220         Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
230         Call Bloquear(toMap, sndIndex, x - 1, y, MapData(Map, x - 1, y).Blocked)
240         Call Bloquear(toMap, sndIndex, x + 1, y, MapData(Map, x + 1, y).Blocked)
            ' Bloqueos inferiores
250         Call Bloquear(toMap, sndIndex, x, y + 1, MapData(Map, x, y + 1).Blocked)
260         Call Bloquear(toMap, sndIndex, x - 1, y + 1, MapData(Map, x - 1, y + 1).Blocked)
270         Call Bloquear(toMap, sndIndex, x + 1, y + 1, MapData(Map, x + 1, y + 1).Blocked)

280     Case 4
            ' Bloqueos superiores
290         Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
            ' Bloqueos inferiores
300         Call Bloquear(toMap, sndIndex, x, y + 1, MapData(Map, x, y + 1).Blocked)

310     Case 5                 'Ver WyroX
            ' Bloqueos vertical ver ReyarB
320         Call Bloquear(toMap, sndIndex, x + 1, y, MapData(Map, x + 1, y).Blocked)
330         Call Bloquear(toMap, sndIndex, x + 1, y - 1, MapData(Map, x + 1, y - 1).Blocked)

            ' Bloqueos horizontal
340         Call Bloquear(toMap, sndIndex, x, y - 2, MapData(Map, x, y - 2).Blocked)
350         Call Bloquear(toMap, sndIndex, x - 1, y - 2, MapData(Map, x - 1, y - 2).Blocked)


360     Case 6                 ' Ver WyroX
            ' Bloqueos superiores ver ReyarB
370         Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
380         Call Bloquear(toMap, sndIndex, x, y - 1, MapData(Map, x, y - 1).Blocked)

            ' Bloqueos inferiores
390         Call Bloquear(toMap, sndIndex, x, y - 2, MapData(Map, x, y - 2).Blocked)
400         Call Bloquear(toMap, sndIndex, x + 1, y - 2, MapData(Map, x + 1, y - 2).Blocked)
410 End Select

420 Exit Sub

MostrarBloqueosPuerta_Err:
430 Call TraceError(Err.Number, Err.Description, "General.MostrarBloqueosPuerta", Erl)

End Sub

Sub BloquearPuerta(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Bloquear As Boolean)

Dim ModPuerta                   As Integer

10  On Error GoTo BloquearPuerta_Err

    'ver reyarb
    With MapData(Map, x, y)

20      ModPuerta = ObjData(.ObjInfo.ObjIndex).Subtipo

30      Select Case ModPuerta

            Case 0             'puerta 2 tiles
                ' Bloqueos superiores
40              .Blocked = IIf(Bloquear, .Blocked Or e_Block.NORTH, .Blocked And Not e_Block.NORTH)
50              MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.NORTH, MapData(Map, x - 1, y).Blocked And Not e_Block.NORTH)

                ' Cambio bloqueos inferiores
60              MapData(Map, x, y + 1).Blocked = IIf(Bloquear, MapData(Map, x, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x, y + 1).Blocked And Not e_Block.SOUTH)
70              MapData(Map, x - 1, y + 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x - 1, y + 1).Blocked And Not e_Block.SOUTH)

80          Case 1
                ' para palancas o teclas sin modicar bloqueos en X,Y

90          Case 2             ' puerta 3 tiles 1 arriba
                ' Bloqueos superiores
100             MapData(Map, x, y - 1).Blocked = IIf(Bloquear, MapData(Map, x, y - 1).Blocked Or e_Block.NORTH, MapData(Map, x, y - 1).Blocked And Not e_Block.NORTH)
110             MapData(Map, x - 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y - 1).Blocked Or e_Block.NORTH, MapData(Map, x - 1, y - 1).Blocked And Not e_Block.NORTH)
120             MapData(Map, x + 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x + 1, y - 1).Blocked Or e_Block.NORTH, MapData(Map, x + 1, y - 1).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
130             .Blocked = IIf(Bloquear, .Blocked Or e_Block.SOUTH, .Blocked And Not e_Block.SOUTH)
140             MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.SOUTH, MapData(Map, x - 1, y).Blocked And Not e_Block.SOUTH)
150             MapData(Map, x + 1, y).Blocked = IIf(Bloquear, MapData(Map, x + 1, y).Blocked Or e_Block.SOUTH, MapData(Map, x + 1, y).Blocked And Not e_Block.SOUTH)

160         Case 3             ' puerta 3 tiles
                ' Bloqueos superiores
170             .Blocked = IIf(Bloquear, .Blocked Or e_Block.NORTH, .Blocked And Not e_Block.NORTH)
180             MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.NORTH, MapData(Map, x - 1, y).Blocked And Not e_Block.NORTH)
190             MapData(Map, x + 1, y).Blocked = IIf(Bloquear, MapData(Map, x + 1, y).Blocked Or e_Block.NORTH, MapData(Map, x + 1, y).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
200             MapData(Map, x, y + 1).Blocked = IIf(Bloquear, MapData(Map, x, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x, y + 1).Blocked And Not e_Block.SOUTH)
210             MapData(Map, x - 1, y + 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x - 1, y + 1).Blocked And Not e_Block.SOUTH)
220             MapData(Map, x + 1, y + 1).Blocked = IIf(Bloquear, MapData(Map, x + 1, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x + 1, y + 1).Blocked And Not e_Block.SOUTH)

230         Case 4             'puerta 1 tiles
                ' Bloqueos superiores
240             .Blocked = IIf(Bloquear, .Blocked Or e_Block.NORTH, .Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
250             MapData(Map, x, y + 1).Blocked = IIf(Bloquear, MapData(Map, x, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x, y + 1).Blocked And Not e_Block.SOUTH)

260         Case 5             'Ver WyroX
                ' Bloqueos  vertical ver ReyarB
270             MapData(Map, x + 1, y).Blocked = IIf(Bloquear, MapData(Map, x + 1, y).Blocked Or e_Block.ALL_SIDES, MapData(Map, x + 1, y).Blocked And Not e_Block.ALL_SIDES)
280             MapData(Map, x + 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x + 1, y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, x + 1, y - 1).Blocked And Not e_Block.ALL_SIDES)

                ' Cambio horizontal
290             MapData(Map, x, y - 2).Blocked = IIf(Bloquear, MapData(Map, x, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x, y - 2).Blocked And Not e_Block.ALL_SIDES)
300             MapData(Map, x - 1, y - 2).Blocked = IIf(Bloquear, MapData(Map, x - 1, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x - 1, y - 2).Blocked And Not e_Block.ALL_SIDES)

310         Case 6             ' Ver Wyrox
                ' Bloqueos vertical ver ReyarB
320             MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.ALL_SIDES, MapData(Map, x - 1, y).Blocked And Not e_Block.ALL_SIDES)
330             MapData(Map, x - 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, x - 1, y - 1).Blocked And Not e_Block.ALL_SIDES)

                ' Cambio bloqueos Puerta abierta
340             MapData(Map, x, y - 2).Blocked = IIf(Bloquear, MapData(Map, x, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x, y - 2).Blocked And Not e_Block.ALL_SIDES)
350             MapData(Map, x + 1, y + 2).Blocked = IIf(Bloquear, MapData(Map, x + 1, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x + 1, y - 2).Blocked And Not e_Block.ALL_SIDES)
360     End Select

    End With

    ' Mostramos a todos
370 Call MostrarBloqueosPuerta(True, Map, x, y)

380 Exit Sub

BloquearPuerta_Err:
390 Call TraceError(Err.Number, Err.Description, "General.BloquearPuerta", Erl)

End Sub

Function HayCosta(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

10  On Error GoTo HayCosta_Err

    'Ladder 10 - 2 - 2010
    'Chequea si hay costa en los tiles proximos al usuario
20  If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And y > 0 And y < 101 Then
30      If ((MapData(Map, x, y).Graphic(1) >= 22552 And MapData(Map, x, y).Graphic(1) <= 22599) Or (MapData(Map, x, y).Graphic(1) >= 7283 And MapData(Map, x, y).Graphic(1) <= 7378) Or (MapData(Map, x, y).Graphic(1) >= 13387 And MapData(Map, x, y).Graphic(1) <= 13482)) And MapData(Map, x, y).Graphic(2) = 0 Then
40          HayCosta = True
50      Else
60          HayCosta = False
70      End If
80  Else
90      HayCosta = False
100 End If

110 Exit Function

HayCosta_Err:
120 Call TraceError(Err.Number, Err.Description, "General.HayCosta", Erl)


End Function

Function HayAgua(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

10  On Error GoTo HayAgua_Err

20  With MapData(Map, x, y)
30      If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And y > 0 And y < 101 Then
40          HayAgua = (.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                      (.Graphic(1) >= 124 And .Graphic(1) <= 139) Or _
                      (.Graphic(1) >= 24223 And .Graphic(1) <= 24238) Or _
                      (.Graphic(1) >= 24303 And .Graphic(1) <= 24318) Or _
                      (.Graphic(1) >= 468 And .Graphic(1) <= 483) Or _
                      (.Graphic(1) >= 44668 And .Graphic(1) <= 44683) Or _
                      (.Graphic(1) >= 24143 And .Graphic(1) <= 24158) Or _
                      (.Graphic(1) >= 12628 And .Graphic(1) <= 12643) Or _
                      (.Graphic(1) >= 2948 And .Graphic(1) <= 2963)
50      Else
60          HayAgua = False
70      End If
80  End With

90  Exit Function

HayAgua_Err:
100 Call TraceError(Err.Number, Err.Description, "General.HayAgua", Erl)


End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean

    On Error GoTo EsArbol_Err

100 EsArbol = GrhIndex = 11905 Or GrhIndex = 644 Or GrhIndex = 1880 Or GrhIndex = 11906 Or GrhIndex = 12160 Or GrhIndex = 6597 Or GrhIndex = 2548 Or GrhIndex = 2549 Or _
              GrhIndex = 15110 Or GrhIndex = 15109 Or GrhIndex = 15108 Or GrhIndex = 11904 Or _
              GrhIndex = 7220 Or GrhIndex = 50990 Or GrhIndex = 55626 Or GrhIndex = 55627 Or GrhIndex = 55630 Or GrhIndex = 55632 Or GrhIndex = 55633 Or _
              GrhIndex = 55635 Or GrhIndex = 55638 Or GrhIndex = 12584 Or GrhIndex = 50985 Or GrhIndex = 15510 Or GrhIndex = 14775 Or GrhIndex = 14687 Or _
              GrhIndex = 11903 Or GrhIndex = 735 Or GrhIndex = 15698 Or GrhIndex = 14504 Or GrhIndex = 15697 Or _
              GrhIndex = 6598 Or GrhIndex = 1121 Or GrhIndex = 1878 Or GrhIndex = 9513 Or GrhIndex = 9514 Or _
              GrhIndex = 9515 Or GrhIndex = 9518 Or GrhIndex = 9519 Or GrhIndex = 9520 Or GrhIndex = 9529

    Exit Function

EsArbol_Err:
102 Call TraceError(Err.Number, Err.Description, "General.EsArbol", Erl)


End Function

Public Function HayLava(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

10  On Error GoTo HayLava_Err

20  If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And y > 0 And y < 101 Then
30      If MapData(Map, x, y).Graphic(1) >= 5837 And MapData(Map, x, y).Graphic(1) <= 5852 Or MapData(Map, x, y).Graphic(1) >= 16101 And MapData(Map, x, y).Graphic(1) <= 16116 Then
40          HayLava = True
50      Else
60          HayLava = False
70      End If
80  Else
90      HayLava = False
100 End If

110 Exit Function

HayLava_Err:
120 Call TraceError(Err.Number, Err.Description, "General.HayLava", Erl)


End Function

Sub ApagarFogatas()

Dim MapaActual                  As Long
Dim y                           As Long
Dim x                           As Long
Dim obj                         As t_Obj

    'Ladder /ApagarFogatas
    On Error GoTo ErrHandler

100 obj.ObjIndex = FOGATA_APAG
102 obj.amount = 1

104 For MapaActual = 1 To NumMaps
106     For y = YMinMapSize To YMaxMapSize
108         For x = XMinMapSize To XMaxMapSize
110             If MapInfo(MapaActual).lluvia Then
112                 If MapData(MapaActual, x, y).ObjInfo.ObjIndex = FOGATA Then
114                     Call EraseObj(MAX_INVENTORY_OBJS, MapaActual, x, y)
116                     Call MakeObj(obj, MapaActual, x, y)
                    End If
                End If
118         Next x
120     Next y
122 Next MapaActual
    Exit Sub

ErrHandler:
124 Call LogError("Error producido al apagar las fogatas de " & x & "-" & y & " del mapa: " & MapaActual & "    -" & Err.Description)

End Sub

Sub LoadMapData()

Dim Map                         As Integer
Dim TempInt                     As Integer
Dim npcfile                     As String
Dim NormalMapsCount             As Integer

10  On Error GoTo man

20  If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

    #If UNIT_TEST = 1 Then
        'We only need 50 maps for unit testing
30      NumMaps = 50
40      Debug.Print "UNIT_TEST Enabled Loading just " & NumMaps & " maps"
    #Else
50      If RunningInVB() Then
            'VB runs out of memory when debugging
60          NumMaps = 300
70      Else
80          NumMaps = CountFiles(MapPath, "*.csm") - 1
90      End If
    #End If

100 NormalMapsCount = NumMaps
110 NumMaps = NormalMapsCount + InstanceMapCount

120 Call InitAreas

130 frmCargando.cargar.Min = 0
140 frmCargando.cargar.max = NormalMapsCount
150 frmCargando.cargar.value = 0
160 frmCargando.ToMapLbl.Visible = True

170 ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As t_MapBlock

180 ReDim MapInfo(1 To NumMaps) As t_MapInfo

190 For Map = 1 To NormalMapsCount
200     frmCargando.ToMapLbl = Map & "/" & NormalMapsCount
210     Call CargarMapaFormatoCSM(Map, MapPath & "Mapa" & Map & ".csm")
220     frmCargando.cargar.value = frmCargando.cargar.value + 1
230     DoEvents
240     Sleep 1
250 Next Map

260 frmCargando.ToMapLbl.Visible = False
270 Call InstanceManager.InitializeInstanceHeap(InstanceMapCount, NormalMapsCount + 1)

280 Exit Sub

man:
290 Call MsgBox("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
300 Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapaFormatoCSM(ByVal Map As Long, ByVal MAPFl As String)

Dim npcfile                     As String
Dim fh                          As Integer
Dim MH                          As t_MapHeader
Dim Blqs()                      As t_DatosBloqueados
Dim L1()                        As t_DatosGrh
Dim L2()                        As t_DatosGrh
Dim L3()                        As t_DatosGrh
Dim L4()                        As t_DatosGrh
Dim Triggers()                  As t_DatosTrigger
Dim Luces()                     As t_DatosLuces
Dim Particulas()                As t_DatosParticulas
Dim Objetos()                   As t_DatosObjs
Dim NPCs()                      As t_DatosNPC
Dim TEs()                       As t_DatosTE
Dim RandomTeleports(MAX_RANDOM_TELEPORT_IN_MAP) As Integer
Dim randomTeleportCount         As Integer
Dim body                        As Integer
Dim head                        As Integer
Dim Heading                     As Byte
Dim SailingTiles                As Long
Dim TotalTiles                  As Long
Dim i                           As Long
Dim j                           As Long
Dim x                           As Integer
Dim y                           As Integer

    On Error GoTo ErrorHandler:

10  randomTeleportCount = 0

20  If Not FileExist(MAPFl, vbNormal) Then
30      Call TraceError(404, "Estas tratando de cargar un MAPA que NO EXISTE" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
40      Exit Sub
50  End If

60  If FileLen(MAPFl) = 0 Then
70      Call TraceError(500, "Se trato de cargar un mapa corrupto o mal generado" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
80      Exit Sub
90  End If

100 fh = FreeFile

110 Open MAPFl For Binary As fh

120 Get #fh, , MH
130 Get #fh, , MapSize
140 Get #fh, , MapDat

    Rem Get #fh, , L1

150 With MH

        'Cargamos Bloqueos
160     If .NumeroBloqueados > 0 Then
170         ReDim Blqs(1 To .NumeroBloqueados)
180         Get #fh, , Blqs
190         For i = 1 To .NumeroBloqueados
200             MapData(Map, Blqs(i).x, Blqs(i).y).Blocked = Blqs(i).Lados
210         Next i
220     End If

        'Cargamos Layer 1

230     If .NumeroLayers(1) > 0 Then
240         ReDim L1(1 To .NumeroLayers(1))
250         Get #fh, , L1
260         For i = 1 To .NumeroLayers(1)
270             x = L1(i).x
280             y = L1(i).y
290             MapData(Map, x, y).Graphic(1) = L1(i).GrhIndex

300             TotalTiles = TotalTiles + 1
310             If HayAgua(Map, x, y) Then
320                 MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_AGUA
330                 SailingTiles = SailingTiles + 1
340             End If
350         Next i
360     End If

        'Cargamos Layer 2
370     If .NumeroLayers(2) > 0 Then
380         ReDim L2(1 To .NumeroLayers(2))
390         Get #fh, , L2
400         For i = 1 To .NumeroLayers(2)
410             x = L2(i).x
420             y = L2(i).y
430             MapData(Map, x, y).Graphic(2) = L2(i).GrhIndex
440             MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked And Not FLAG_AGUA
450         Next i

460     End If

470     If .NumeroLayers(3) > 0 Then
480         ReDim L3(1 To .NumeroLayers(3))
490         Get #fh, , L3

500         For i = 1 To .NumeroLayers(3)
510             x = L3(i).x
520             y = L3(i).y
530             MapData(Map, x, y).Graphic(3) = L3(i).GrhIndex

540             If EsArbol(L3(i).GrhIndex) Then
550                 MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_ARBOL
560             End If
570         Next i

580     End If

590     If .NumeroLayers(4) > 0 Then
600         ReDim L4(1 To .NumeroLayers(4))
610         Get #fh, , L4
620         For i = 1 To .NumeroLayers(4)
630             MapData(Map, L4(i).x, L4(i).y).Graphic(4) = L4(i).GrhIndex
640         Next i

650     End If

660     If .NumeroTriggers > 0 Then
670         ReDim Triggers(1 To .NumeroTriggers)
680         Get #fh, , Triggers

690         For i = 1 To .NumeroTriggers
700             x = Triggers(i).x
710             y = Triggers(i).y
720             MapData(Map, x, y).trigger = Triggers(i).trigger

                ' Trigger detalles en agua
730             If Triggers(i).trigger = e_Trigger.DETALLEAGUA Then
                    ' Vuelvo a poner flag agua
740                 MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_AGUA
750             End If

760             If Triggers(i).trigger = e_Trigger.VALIDONADO Or Triggers(i).trigger = e_Trigger.NADOCOMBINADO Or Triggers(i).trigger = e_Trigger.NADOBAJOTECHO Then
                    ' Vuelvo a poner flag agua
770                 MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_AGUA
780             End If
790         Next i
800     End If

810     If .NumeroParticulas > 0 Then
820         ReDim Particulas(1 To .NumeroParticulas)
830         Get #fh, , Particulas

840         For i = 1 To .NumeroParticulas
850             MapData(Map, Particulas(i).x, Particulas(i).y).ParticulaIndex = Particulas(i).Particula
860             MapData(Map, Particulas(i).x, Particulas(i).y).ParticulaIndex = 0
870         Next i
880     End If

890     If .NumeroLuces > 0 Then
900         ReDim Luces(1 To .NumeroLuces)
910         Get #fh, , Luces

920         For i = 1 To .NumeroLuces
930             MapData(Map, Luces(i).x, Luces(i).y).Luz.Color = Luces(i).Color
940             MapData(Map, Luces(i).x, Luces(i).y).Luz.Rango = Luces(i).Rango
950             MapData(Map, Luces(i).x, Luces(i).y).Luz.Color = 0
960             MapData(Map, Luces(i).x, Luces(i).y).Luz.Rango = 0
970         Next i
980     End If

990     If .NumeroOBJs > 0 Then
1000        ReDim Objetos(1 To .NumeroOBJs)
1010        Get #fh, , Objetos
1020        For i = 1 To .NumeroOBJs
1030            MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.ObjIndex = Objetos(i).ObjIndex
1040            With ObjData(Objetos(i).ObjIndex)
1050                Select Case .OBJType
                        Case e_OBJType.otOreDeposit, e_OBJType.otTrees
1060                        MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.amount = ObjData(Objetos(i).ObjIndex).VidaUtil
1070                        MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.data = &H7FFFFFFF    ' Ultimo uso = Max Long
1080                    Case Else
1090                        MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.amount = Objetos(i).ObjAmmount
1100                End Select
1110                If .OBJType = otTeleport And .Subtipo = e_TeleportSubType.eTransportNetwork Then
1120                    RandomTeleports(randomTeleportCount) = i
1130                    randomTeleportCount = randomTeleportCount + 1
1140                End If
1150            End With
1160        Next i
1170    End If

1180    If .NumeroNPCs > 0 Then
1190        ReDim NPCs(1 To .NumeroNPCs)
1200        Get #fh, , NPCs

            Dim NumNpc As Integer, NpcIndex As Integer

1210        For i = 1 To .NumeroNPCs
1220            NumNpc = NPCs(i).NpcIndex
1230            If NumNpc > 0 Then
1240                npcfile = DatPath & "NPCs.dat"
1250                NpcIndex = OpenNPC(NumNpc)

1260                If NpcIndex > 0 Then
1270                    MapData(Map, NPCs(i).x, NPCs(i).y).NpcIndex = NpcIndex
1280                    NpcList(NpcIndex).pos.Map = Map
1290                    NpcList(NpcIndex).pos.x = NPCs(i).x
1300                    NpcList(NpcIndex).pos.y = NPCs(i).y
                        '  guardo siempre la pos original... puede sernos útil ;)
1310                    NpcList(NpcIndex).Orig = NpcList(NpcIndex).pos

1320                    If LenB(NpcList(NpcIndex).name) = 0 Then
1330                        MapData(Map, NPCs(i).x, NPCs(i).y).NpcIndex = 0
1340                    Else
1350                        Call MakeNPCChar(True, 0, NpcIndex, Map, NPCs(i).x, NPCs(i).y)
1360                    End If
1370                Else
                        ' Lo guardo en los logs + aparece en el Debug.Print
1380                    Call TraceError(404, "NPC no existe en los .DAT's o está mal dateado. Posicion: " & Map & "-" & NPCs(i).x & "-" & NPCs(i).y, "ES.CargarMapaFormatoCSM")
1390                End If
1400            End If
1410        Next i
1420    End If

1430    If .NumeroTE > 0 Then
1440        ReDim TEs(1 To .NumeroTE)
1450        Get #fh, , TEs

1460        For i = 1 To .NumeroTE
1470            MapData(Map, TEs(i).x, TEs(i).y).TileExit.Map = TEs(i).DestM
1480            MapData(Map, TEs(i).x, TEs(i).y).TileExit.x = TEs(i).DestX
1490            MapData(Map, TEs(i).x, TEs(i).y).TileExit.y = TEs(i).DestY
1500        Next i
1510    End If
1520 End With
1530 Close fh

    '  Nuevo sistema de restricciones
1540 If Not IsNumeric(MapDat.restrict_mode) Then
        ' Solo se usaba el "NEWBIE"
1550    If UCase$(MapDat.restrict_mode) = "NEWBIE" Then
1560        MapDat.restrict_mode = "1"
1570    Else
1580        MapDat.restrict_mode = "0"
1590    End If
1600 End If

1610 If SailingTiles * 100 / TotalTiles > SvrConfig.GetValue("FISHING_REQUIRED_PERCENT") And Not MapDat.Seguro Then
1620    Call AddFishingPoolsToMap(Map)
1630 End If


1640 With MapInfo(Map)

1650    .map_name = MapDat.map_name
1660    .MapResource = Map
1670    .ambient = MapDat.ambient
1680    .backup_mode = MapDat.backup_mode
1690    .base_light = MapDat.base_light
1700    .Newbie = (val(MapDat.restrict_mode) And 1) <> 0
1710    .SinMagia = (val(MapDat.restrict_mode) And 2) <> 0
1720    .NoPKs = (val(MapDat.restrict_mode) And 4) <> 0
1730    .NoCiudadanos = (val(MapDat.restrict_mode) And 8) <> 0
1740    .SinInviOcul = (val(MapDat.restrict_mode) And 16) <> 0
1750    .SoloClanes = (val(MapDat.restrict_mode) And 32) <> 0
1760    .NoMascotas = (val(MapDat.restrict_mode) And 64) <> 0
1770    .OnlyGroups = (val(MapDat.restrict_mode) And 128) <> 0
1780    .OnlyPatreon = (val(MapDat.restrict_mode) And 256) <> 0
1790    .ResuCiudad = val(GetVar(DatPath & "Map.dat", "RESUCIUDAD", Map)) <> 0
1800    .letter_grh = MapDat.letter_grh
1810    .lluvia = MapDat.lluvia
1820    .music_numberHi = MapDat.music_numberHi
1830    .music_numberLow = MapDat.music_numberLow
1840    .niebla = MapDat.niebla
1850    .Nieve = MapDat.Nieve
1860    .MinLevel = MapDat.level And &HFF
1870    .MaxLevel = (MapDat.level And &HFF00) / &H100
1880    .Seguro = MapDat.Seguro
1890    .terrain = MapDat.terrain
1900    .zone = MapDat.zone
1910    .DropItems = True

1920    If EsMapaNoDrop(Map) Then
1930        .DropItems = False
1940    End If

1950    .FriendlyFire = True
1960    .KeepInviOnAttack = val(GetVar(DatPath & "Map.dat", "KeepInviOnAttack", Map)) <> 0
1970    .ForceUpdate = val(GetVar(DatPath & "Map.dat", "ForceUpdateAi", Map)) <> 0

1980    If LenB(MapDat.Salida) <> 0 Then
            Dim Fields()        As String
1990        Fields = Split(MapDat.Salida, "-")
2000        .Salida.Map = val(Fields(0))
2010        .Salida.x = val(Fields(1))
2020        .Salida.y = val(Fields(2))
2030    End If

2040    If randomTeleportCount > 0 Then
2050        ReDim .TransportNetwork(randomTeleportCount - 1) As t_TransportNetworkExit
2060        For i = 0 To randomTeleportCount - 1
2070            .TransportNetwork(i).TileX = Objetos(RandomTeleports(i)).x
2080            .TransportNetwork(i).TileY = Objetos(RandomTeleports(i)).y
2090        Next i
2100    End If
2110 End With

2120 Exit Sub

ErrorHandler:
2130 Close fh
2140 Call TraceError(Err.Number, Err.Description, "ES.CargarMapaFormatoCSM", Erl)

End Sub

Sub AddFishingPoolsToMap(ByVal Map As Integer)
    Dim i As Integer
    For i = 1 To SvrConfig.GetValue("FISHING_TILES_ON_MAP")
        Call CreateFishingPool(Map)
    Next i
End Sub

Public Sub CreateFishingPool(ByVal Map As Integer)
    Dim x, y As Integer
    Do
        x = RandomNumber(12, 88)
        y = RandomNumber(12, 88)
    Loop While MapData(Map, x, y).ObjInfo.ObjIndex <> 0 Or Not HayAgua(Map, x, y)
    MapData(Map, x, y).ObjInfo.ObjIndex = SvrConfig.GetValue("FISHING_POOL_ID")
    MapData(Map, x, y).ObjInfo.amount = ObjData(SvrConfig.GetValue("FISHING_POOL_ID")).VidaUtil
    MapData(Map, x, y).ObjInfo.data = &H7FFFFFFF ' Ultimo uso = Max Long
End Sub

Function IsUserAtPos(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte) As Boolean
    IsUserAtPos = MapData(Map, x, y).UserIndex > 0

End Function

Function IsNpcAtPos(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte)
    IsNpcAtPos = MapData(Map, x, y).NpcIndex > 0

End Function
