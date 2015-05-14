Attribute VB_Name = "modLunarCompress"
'Visual Basic module for Lunar Compress
' --> updated February 1, 2003 for LC version 1.40

'Module by Bouche
'Contact Details:
'E-mail: andrewrl87@yahoo.com
'AIM: BoucheanBouche
'MSN: andrewrl87@yahoo.com
'ICQ: 69427310
'Y!: andrewrl87
'AcmlmBoard ID: 63
'WWW: bouche.kafuka.org
'IRC: #romhacking on irc.hexnet.com

'Compression Formats
Public Enum LunarCompressionMode
 LC_LZ1 = 0
 LC_LZ2 = 1
 LC_LZ3 = 2
 LC_LZ4 = 3
 LC_LZ5 = 4
 LC_LZ6 = 5
 LC_LZ7 = 6
 LC_LZ8 = 7
 LC_LZ9 = 8
 LC_LZ10 = 9
 LC_LZ11 = 10
 LC_LZ12 = 11
 LC_LZ13 = 12
 LC_RLE1 = 100
 LC_RLE2 = 101
 LC_RLE3 = 102
End Enum

Public Enum LunarExpansionMode
 LC_48_EXHIROM = 48
 LC_48_EXHIROM_1 = 304    'Higher compatibility, but uses up to 1 meg of the new space.  Do not use this unless the ROM doesn't load or has problems with the other options.
 LC_64_EXHIROM = 64
 LC_64_EXHIROM_1 = 320    'Higher compatibility, but uses up to 2 meg of the new space.  Do not use this unless the ROM doesn't load or has problems with the other options.
 LC_48_EXLOROM_1 = 4144   'For LoROMs that use the 00:8000-6F:FFFF
 LC_48_EXLOROM_2 = 8240   'For LoROMs that use the 80:8000-FF:FFFF map.
 LC_48_EXLOROM_3 = 16432  'Higher compatibility, but uses up most of the new space.  Do not use this unless the ROM doesn't load or has problems with the other options.
 LC_64_EXLOROM_1 = 4160   'For LoROMs that use the 00:8000-6F:FFFF
 LC_64_EXLOROM_2 = 8256   'For LoROMs that use the 80:8000-FF:FFFF map.
 LC_64_EXLOROM_3 = 16448  'Higher compatibility, but uses up most of the new space.  Do not use this unless the ROM doesn't load or has problems with the other options.
End Enum

Public Enum LunarFileMode
 LC_READONLY = 0
 LC_READWRITE = 1
 LC_CREATEREADWRITE = 2
 LC_LOCKARRAYSIZE = 4
 LC_LOCKARRAYSIZE_2 = 8
 LC_CREATEARRAY = 16
 LC_SAVEONCLOSE = 32
End Enum

Public Enum LunarSeekMode
 LC_NOSEEK = 0
 LC_SEEK = 1
End Enum

Public Enum LunarAddressMode
 LC_NOBANK = 0
 LC_LOROM = 1     'LoROM
 LC_HIROM = 2     'HiROM
 LC_EXHIROM = 4   'Extended HiROM
 LC_EXLOROM = 8   'Extended LoROM
 LC_LOROM_2 = 16  'LoROM, always converts to 80:8000 map
 LC_EXROM = 4     'same as LC_EXHIROM (depreciated)
End Enum

Public Enum LunarHeaderMode
 LC_NOHEADER = 0
 LC_HEADER = 1
End Enum

Public Enum LunarIPSFlags
 LC_IPSLOG = &H80000000
 LC_IPSQUIET = &H40000000
End Enum

Public Enum LunarGraphicsMode
 LC_1BPP = 1
 LC_2BPP = 2
 LC_3BPP = 3
 LC_4BPP = 4
 LC_5BPP = 5
 LC_6BPP = 6
 LC_7BPP = 7
 LC_8BPP = 8
 LC_4BPP_GBA = &H14   'unofficial support
End Enum

Public Enum LunarRenderFlags
 LC_INVERT_TRANSPARENT = 1
 LC_INVERT_OPAQUE = 2
 LC_INVERT = 3
 LC_RED_TRANSPARENT = 4
 LC_RED_OPAQUE = 8
 LC_RED = 12
 LC_GREEN_TRANSPARENT = 16
 LC_GREEN_OPAQUE = 32
 LC_GREEN = 48
 LC_BLUE_TRANSPARENT = 64
 LC_BLUE_OPAQUE = 128
 LC_BLUE = 192
 LC_TRANSLUCENT = 256
 LC_HALF_COLOR = 512    'half-color mode
 LC_SCREEN_ADD = 1024   'sub-screen addition
 LC_SCREEN_SUB = 2048   'sub-screen subtraction
 LC_PRIORITY_0 = 4096
 LC_PRIORITY_1 = 8192
 LC_PRIORITY_2 = 16384
 LC_PRIORITY_3 = 32768
 LC_DRAW = 61440
 LC_OPAQUE = 65536
 LC_SPRITE = 131072
 LC_SPRITE_TRANSLUCENT = 262144
 LC_2BPP_GFX = &H80000
 LC_TILE_16 = &H100000
 LC_TILE_32 = &H200000
 LC_TILE_64 = &H400000
End Enum

Public Enum LunarRATFlags
 RATF_FORMAT = &HFF         'bits reserved to specify LC compressed format (DO NOT USE THIS VALUE AS A FLAG!)
 RATF_LOROM = &H100         'use LoROM banks
 RATF_HIROM = &H200         'use HiROM banks
 RATF_EXLOROM = &H10000     'NOT same as RATF_LOROM
 RATF_EXHIROM = &H400       'NOT same as RATF_HIROM
 RATF_EXROM = &H400         'same as RATF_EXHIROM (old)
 RATF_COMPRESSED = &H800    'data to erase is compressed; can decompress to get size using LC format specified
 RATF_NOERASERAT = &H1000   'don't erase RAT tag
 RATF_NOWRITERAT = &H2000   'don't write RAT tag
 RATF_NOERASEDATA = &H4000  'don't erase user data
 RATF_NOWRITEDATA = &H8000  'don't write user data
End Enum

'Lunar Compress Version
Declare Function LunarVersion Lib "lunar compress.dll" () As Long

'File Opening/Closing
Declare Function LunarOpenFile Lib "lunar compress.dll" (ByVal filename As String, ByVal FileMode As LunarFileMode) As Boolean
Declare Function LunarOpenRAMFile Lib "lunar compress.dll" (ByVal data As Long, ByVal FileMode As LunarFileMode, ByVal size As Long)
Declare Function LunarSaveRAMFile Lib "lunar compress.dll" (ByVal filename As String)
Declare Function LunarCloseFile Lib "lunar compress.dll" () As Boolean

'Retrieving File Size
Declare Function LunarGetFileSize Lib "lunar compress.dll" () As Long

'Reading/Writing
Declare Function LunarReadFile Lib "lunar compress.dll" (destination As Byte, ByVal size As Long, ByVal Address As Long, ByVal seekx As LunarSeekMode) As Long
Declare Function LunarWriteFile Lib "lunar compress.dll" (source As Byte, ByVal size As Long, ByVal Address As Long, ByVal seekx As LunarSeekMode) As Long

'Address Conversion
Declare Function LunarSNEStoPC Lib "lunar compress.dll" (ByVal Pointer As Long, ByVal ROMType As LunarAddressMode, ByVal Header As LunarHeaderMode) As Long
Declare Function LunarPCtoSNES Lib "lunar compress.dll" (ByVal Pointer As Long, ByVal ROMType As LunarAddressMode, ByVal Header As LunarHeaderMode) As Long

'Compression
Declare Function LunarDecompress Lib "lunar compress.dll" (destination As Byte, ByVal AddressToStart As Long, ByVal MaxDataSize As Long, ByVal Format As Long, ByVal Format2 As Long, LastRomPosition As Long) As Long
Declare Function LunarRecompress Lib "lunar compress.dll" (source As Byte, destination As Byte, ByVal DataSize As Long, ByVal MaxDataSize As Long, ByVal Format As LunarCompressionMode, ByVal Format2 As Long) As Long

'ROM Space and Area
Declare Function LunarEraseArea Lib "lunar compress.dll" (ByVal Address As Long, ByVal size As Long) As Boolean
Declare Function LunarVerifyFreeSpace Lib "lunar compress.dll" (ByVal addressstart As Long, ByVal addressend As Long, ByVal size As Long, ByVal BankType As LunarAddressMode) As Long
Declare Function LunarExpandROM Lib "lunar compress.dll" (ByVal mbits As Long) As Long

'IPS Functions
Declare Function LunarIPSCreate Lib "lunar compress.dll" (Optional ByVal hWnd As Long = 0, Optional ByVal IPSFileName As String = "", Optional ByVal ROMFileName As String = "", Optional ByVal ROM2FileName As String = "", Optional ByVal IPSFlags As LunarIPSFlags) As Boolean
Declare Function LunarIPSApply Lib "lunar compress.dll" (Optional ByVal hWnd As Long = 0, Optional ByVal IPSFileName As String = "", Optional ByVal ROMFileName As String = "", Optional ByVal IPSFlags As LunarIPSFlags = 0) As Boolean

'Pixel Map and BPP Map
Declare Function LunarCreatePixelMap Lib "lunar compress.dll" (source As Byte, destination As Byte, ByVal numtiles As Long, ByVal gfxtype As LunarGraphicsMode) As Boolean
Declare Function LunarCreateBPPMap Lib "lunar compress.dll" (source As Byte, destination As Byte, ByVal numtiles As Long, ByVal gfxtype As LunarGraphicsMode) As Boolean

'Palette Conversion
Declare Function LunarSNEStoPCRGB Lib "lunar compress.dll" (ByVal snesolor As Long) As Long
Declare Function LunarPCtoSNESRGB Lib "lunar compress.dll" (ByVal pccolor As Long) As Long

'Rendering
Declare Function LunarRender8x8 Lib "lunar compress.dll" (ByVal mapBits As Long, ByVal Width As Long, ByVal Height As Long, ByVal displayX As Long, ByVal displayY As Long, pixelMap As Byte, PCPalette As Long, ByVal map8Tile As Long, ByVal extra As LunarRenderFlags) As Long
 
 'Note: LunarRender8x8 does not work in VB for some reason
 'please contact Bouche if you find a way around this please

'ROM Allocation Tag System
Declare Function LunarWriteRatArea Lib "lunar compress.dll" (ByVal TheData As Long, ByVal size As Long, ByVal PreferredAddress As Long, ByVal MinRange As Long, ByVal MaxRange As Long, ByVal Flags As LunarRATFlags) As Long
Declare Function LunarEraseRatArea Lib "lunar compress.dll" (ByVal Address As Long, ByVal size As Long, ByVal Flags As LunarRATFlags) As Long
Declare Function LunarGetRatAreaSize Lib "lunar compress.dll" (ByVal Address As Long, ByVal Flags As LunarRATFlags) As Long
