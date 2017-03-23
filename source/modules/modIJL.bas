Option Compare Database
Option Explicit
Rem
Rem
Rem
Rem               INTEL CORPORATION PROPRIETARY INFORMATION
Rem  This software is supplied under the terms of a license agreement or
Rem  nondisclosure agreement with Intel Corporation and may not be copied
Rem  or disclosed except in accordance with the terms of that agreement.
Rem      Copyright (c) 1998 Intel Corporation. All Rights Reserved.
Rem
Rem
Rem  File:
Rem    ijl.bas
Rem
Rem  Purpose:
Rem    Intel(R) JPEG Library Visual Basic interface module
Rem
Rem  Version:
Rem    1.2
Rem
Rem




Global Const JBUFSIZE  As Long = 4096




Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        IJLibVersion
Rem
Rem  Purpose:     Stores library version info.
Rem
Rem  Context:
Rem
Rem  Example:
Rem   major           - 1
Rem   minor           - 0
Rem   build           - 1
Rem   Name            - "ijl10.dll"
Rem   Version         - "1.0.1 Beta 1"
Rem   InternalVersion - "1.0.1.1"
Rem   BuildDate       - "Sep 22 1998"
Rem   CallConv        - "DLL"
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type IJLibVersion
  Major           As Long
  Minor           As Long
  build           As Long
  name            As Long 'pointer to C-style string
  Version         As Long 'pointer to C-style string
  InternalVersion As Long 'pointer to C-style string
  BuildDate       As Long 'pointer to C-style string
  CallConv        As Long 'pointer to C-style string
End Type


Rem/*D*
Rem////////////////////////////////////////////////////////////////////////////
Rem// Name:        IJL_RECT
Rem//
Rem// Purpose:     Keep coordinates for rectangle region of image
Rem//
Rem// Context:     Used to specify roi
Rem//
Rem// Fields:
Rem//
Rem////////////////////////////////////////////////////////////////////////////
Rem*D*/

Public Type IJL_RECT
  Left   As Long
  Top    As Long
  right  As Long
  Bottom As Long
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:       IJLIOTYPE
Rem
Rem  Purpose:    Possible types of data read/write/other operations to be
Rem              performed by the functions IJL_Read and IJL_Write.
Rem
Rem              See the Developer's Guide for details on appropriate usage.
Rem
Rem  Fields:
Rem
Rem    IJL_JFILE_XXXXXXX   Indicates JPEG data in a stdio file.
Rem
Rem    IJL_JBUFF_XXXXXXX   Indicates JPEG data in an addressable buffer.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum IJLIOTYPE
  Public Const IJL_SETUP = -1
    
  ' Read JPEG parameters (i.e., height, width, channels, sampling, etc.)
  ' from a JPEG bit stream.
  Public Const IJL_JFILE_READPARAMS = 0
  Public Const IJL_JBUFF_READPARAMS = 1
    
  ' Read a JPEG Interchange Format image.
  Public Const IJL_JFILE_READWHOLEIMAGE = 2
  Public Const IJL_JBUFF_READWHOLEIMAGE = 3
    
  ' Read JPEG tables from a JPEG Abbreviated Format bit stream.
  Public Const IJL_JFILE_READHEADER = 4
  Public Const IJL_JBUFF_READHEADER = 5
    
  ' Read image info from a JPEG Abbreviated Format bit stream.
  Public Const IJL_JFILE_READENTROPY = 6
  Public Const IJL_JBUFF_READENTROPY = 7
    
  ' Write an entire JFIF bit stream.
  Public Const IJL_JFILE_WRITEWHOLEIMAGE = 8
  Public Const IJL_JBUFF_WRITEWHOLEIMAGE = 9
    
  ' Write a JPEG Abbreviated Format bit stream.
  Public Const IJL_JFILE_WRITEHEADER = 10
  Public Const IJL_JBUFF_WRITEHEADER = 11
    
  ' Write image info to a JPEG Abbreviated Format bit stream.
  Public Const IJL_JFILE_WRITEENTROPY = 12
  Private Const IJL_JBUFF_WRITEENTROPY = 13
    
  ' Scaled Decoding Options:

  ' Reads a JPEG image scaled to 1/2 size.
  Public Const IJL_JFILE_READONEHALF = 14
  Public Const IJL_JBUFF_READONEHALF = 15
    
  ' Reads a JPEG image scaled to 1/4 size.
  Public Const IJL_JFILE_READONEQUARTER = 16
  Public Const IJL_JBUFF_READONEQUARTER = 17
    
  ' Reads a JPEG image scaled to 1/8 size.
  Public Const IJL_JFILE_READONEEIGHTH = 18
  Public Const IJL_JBUFF_READONEEIGHTH = 19
    
  ' Reads an embedded thumbnail from a JFIF bit stream.
  Public Const IJL_JFILE_READTHUMBNAIL = 20
  Public Const IJL_JBUFF_READTHUMBNAIL = 21
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        IJL_COLOR
Rem
Rem  Purpose:     Possible color space formats.
Rem
Rem  Note these formats do *not* necessarily denote
Rem  the number of channels in the color space.
Rem  There exists separate "channel" fields in the
Rem  JPEG_CORE_PROPERTIES data structure specifically
Rem  for indicating the number of channels in the
Rem  JPEG and/or DIB color spaces.
Rem
Rem  See the Developer's Guide for details on appropriate usage.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum IJL_COLOR
  Public Const IJL_RGB = 1        ' Red-Green-Blue color space.
  Public Const IJL_BGR = 2        ' Reversed channel ordering from IJL_RGB.
  Public Const IJL_YCBCR = 3      ' Luminance-Chrominance color space as defined
                     ' by CCIR Recommendation 601.
  Public Const IJL_G = 4          ' Grayscale color space.
  Public Const IJL_RGBA_FPX = 5   ' FlashPix RGB 4 channel color space that
                     ' has pre-multiplied opacity.
  Public Const IJL_YCBCRA_FPX = 6 ' FlashPix YCbCr 4 channel color space that
                     ' has pre-multiplied opacity.
  Public Const IJL_OTHER = 255    ' Some other color space not defined by the IJL.
                     ' (This means no color space conversion will be
                     ' done by the IJL.)
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        IJL_JPGSUBSAMPLING
Rem
Rem  Purpose:     Possible subsampling formats used in the JPEG.
Rem
Rem               See the Developer's Guide for details on appropriate usage.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum IJL_JPGSUBSAMPLING
Public Const IJL_NONE = 0    ' Corresponds to "No Subsampling".
                ' Valid on a JPEG w/ any number of channels.
 Public Const IJL_411 = 1    ' Valid on a JPEG w/ 3 channels.
 Public Const IJL_422 = 2    ' Valid on a JPEG w/ 3 channels.
  
 Public Const IJL_4114 = 3   ' Valid on a JPEG w/ 4 channels.
  Public Const IJL_4224 = 4  ' Valid on a JPEG w/ 4 channels.
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        IJL_DIBSUBSAMPLING
Rem
Rem  Purpose:     Possible subsampling formats used in the DIB.
Rem
Rem  See the Developer's Guide for details on appropriate usage.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum IJL_DIBSUBSAMPLING
  'Public Const IJL_NONE = 0  ' Corresponds to "No Subsampling".
                ' Valid on a DIB w/ any number of channels.
  'Public Const IJL_422 = 2   ' Valid on a DIB with YCbYCr color.
  
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        HUFFMAN_TABLE
Rem
Rem  Purpose:     Stores Huffman table information in a fast-to-use format.
Rem
Rem  Context:     Used by Huffman encoder/decoder to access Huffman table
Rem               data.  Raw Huffman tables are formatted to fit this
Rem               structure prior to use.
Rem
Rem  Fields:
Rem    huff_class  0 == DC Huffman or lossless table, 1 == AC table.
Rem    ident       Huffman table identifier, 0-3 valid (Extended Baseline).
Rem    huffelem    Huffman elements for codes <= 8 bits long;
Rem                contains both zero run-length and symbol length in bits.
Rem    huffval     Huffman values for codes 9-16 bits in length.
Rem    mincode     Smallest Huffman code of length n.
Rem    maxcode     Largest Huffman code of length n.
Rem    valptr      Starting index into huffval[] for symbols of length k.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type HUFFMAN_TABLE
  huff_class         As Long
  ident              As Long
  huffelem(0 To 255) As Long
  huffval(0 To 255)  As Integer
  mincode(0 To 16)   As Integer
  maxcode(0 To 17)   As Integer
  valptr(0 To 16)    As Integer
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        JPEGHuffTable
Rem
Rem  Purpose:     Stores pointers to JPEG-binary spec compliant
Rem               Huffman table information.
Rem
Rem  Context:     Used by interface and table methods to specify encoder
Rem               tables to generate and store JPEG images.
Rem
Rem  Fields:
Rem    bits        Points to number of codes of length i (<=16 supported).
Rem    vals        Value associated with each Huffman code.
Rem    hclass      0 == DC table, 1 == AC table.
Rem    ident       Specifies the identifier for this table.
Rem                0-3 for extended JPEG compliance.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type JPEGHuffTable
  bits   As Long
  vals   As Long
  hclass As Byte
  ident  As Byte
  ' IJL use 8 byte pack structures
  pad0   As Byte
  pad1   As Byte
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        QUANT_TABLE
Rem
Rem  Purpose:     Stores quantization table information in a
Rem               fast-to-use format.
Rem
Rem  Context:     Used by quantizer/dequantizer to store formatted
Rem               quantization tables.
Rem
Rem  Fields:
Rem    precision   0 => elements contains 8-bit elements,
Rem                1 => elements contains 16-bit elements.
Rem    ident       Table identifier (0-3).
Rem    elements    Pointer to 64 table elements + 16 extra elements to catch
Rem                input data errors that may cause malfunction of the
Rem                Huffman decoder.
Rem    elarray     Space for elements (see above) plus 8 bytes to align
Rem                to a quadword boundary.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type QUANT_TABLE
  precision        As Long
  ident            As Long
  elements         As Long
  elarray(0 To 83) As Integer
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        JPEGQuantTable
Rem
Rem  Purpose:     Stores pointers to JPEG binary spec compliant
Rem               quantization table information.
Rem
Rem  Context:     Used by interface and table methods to specify encoder
Rem               tables to generate and store JPEG images.
Rem
Rem  Fields:
Rem    quantizer   Zig-zag order elements specifying quantization factors.
Rem    ident       Specifies identifier for this table.
Rem                0-3 valid for Extended Baseline JPEG compliance.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type JPEGQuantTable
  quantizer As Long
  ident     As Byte
  ' IJL use 8 byte pack structures
  pad0      As Byte
  pad1      As Byte
  pad2      As Byte
End Type


Rem
Rem///////////////////////////////////////////////////////////////////////////
Rem  Name:        FRAME_COMPONENT
Rem
Rem  Purpose:     One frame-component structure is allocated per component
Rem               in a frame.
Rem
Rem  Context:     Used by Huffman decoder to manage components.
Rem
Rem  Fields:
Rem    ident       Component identifier.  The tables use this ident to
Rem                determine the correct table for each component.
Rem    hsampling   Horizontal subsampling factor for this component,
Rem                1-4 are legal.
Rem    vsampling   Vertical subsampling factor for this component,
Rem                1-4 are legal.
Rem    quant_sel   Quantization table selector.  The quantization table
Rem                used by this component is determined via this selector.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type FRAME_COMPONENT
  ident     As Long
  hsampling As Long
  vsampling As Long
  quant_sel As Long
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        FRAME
Rem
Rem  Purpose:     Stores frame-specific data.
Rem
Rem  Context:     One Frame structure per image.
Rem
Rem  Fields:
Rem    precision       Sample precision in bits.
Rem    width           Width of the source image in pixels.
Rem    height          Height of the source image in pixels.
Rem    MCUheight       Height of a frame MCU.
Rem    MCUwidth        Width of a frame MCU.
Rem    max_hsampling   Max horiz sampling ratio of any component in the frame.
Rem    max_vsampling   Max vert sampling ratio of any component in the frame.
Rem    ncomps          Number of components/channels in the frame.
Rem    horMCU          Number of horizontal MCUs in the frame.
Rem    totalMCU        Total number of MCUs in the frame.
Rem    comps           Array of 'ncomps' component descriptors.
Rem    restart_interv  Indicates number of MCUs after which to restart the
Rem                    entropy parameters.
Rem    SeenAllDCScans  Used when decoding Multiscan images to determine if
Rem                    all channels of an image have been decoded.
Rem    SeenAllACScans  (See SeenAllDCScans)
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type Frame
  precision      As Long
  width          As Long
  Height         As Long
  MCUheight      As Long
  MCUwidth       As Long
  max_hsampling  As Long
  max_vsampling  As Long
  ncomps         As Long
  horMCU         As Long
  totalMCU       As Long
  comps          As Long
  restart_interv As Long
  SeenAllDCScans As Long
  SeenAllACScans As Long
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        SCAN_COMPONENT
Rem
Rem  Purpose:     One scan-component structure is allocated per component
Rem               of each scan in a frame.
Rem
Rem  Context:     Used by Huffman decoder to manage components within scans.
Rem
Rem  Fields:
Rem    comp        Component number, index to the comps member of FRAME.
Rem    hsampling   Horizontal sampling factor.
Rem    vsampling   Vertical sampling factor.
Rem    dc_table    DC Huffman table pointer for this scan.
Rem    ac_table    AC Huffman table pointer for this scan.
Rem    quant_table Quantization table pointer for this scan.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type SCAN_COMPONENT
  comp       As Long
  hsampling  As Long
  vsampling  As Long
  dc_table   As Long
  ac_table   As Long
  quantTable As Long
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        SCAN
Rem
Rem  Purpose:     One SCAN structure is allocated per scan in a frame.
Rem
Rem  Context:     Used by Huffman decoder to manage scans.
Rem
Rem  Fields:
Rem    ncomps          Number of image components in a scan, 1-4 legal.
Rem    gray_scale      If TRUE, decode only the Y channel.
Rem    start_spec      Start coefficient of spectral or predictor selector.
Rem    end_spec        End coefficient of spectral selector.
Rem    approx_high     High bit position in successive approximation
Rem                    Progressive coding.
Rem    approx_low      Low bit position in successive approximation
Rem                    Progressive coding.
Rem    restart_interv  Restart interval, 0 if disabled.
Rem    curxMCU         Next horizontal MCU index to be processed after
Rem                    an interrupted SCAN.
Rem    curyMCU         Next vertical MCU index to be processed after
Rem                    an interrupted SCAN.
Rem    dc_diff         Array of DC predictor values for DPCM modes.
Rem    comps           Array of ncomps SCAN_COMPONENT component identifiers.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type SCAN
  ncomps          As Long
  gray_scale      As Long
  start_spec      As Long
  end_spec        As Long
  approx_high     As Long
  approx_low      As Long
  restart_interv  As Long
  curxMCU         As Long
  curyMCU         As Long
  dc_diff(0 To 3) As Long
  comps           As Long
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        DCTTYPE
Rem
Rem  Purpose:     Possible algorithms to be used to perform the discrete
Rem               cosine transform (DCT).
Rem
Rem  Fields:
Rem    IJL_AAN     The AAN (Arai, Agui, and Nakajima) algorithm from
Rem                Trans. IEICE, vol. E 71(11), 1095-1097, Nov. 1988.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum DCTTYPE
 Public Const IJL_AAN = 0
  Public Const IJL_IPP = 1
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem Name:        UPSAMPLING_TYPE
Rem
Rem Purpose:            -  Possible algorithms to be used to perform upsampling
Rem
Rem Fields:
Rem  IJL_BOX_FILTER      - the algorithm is simple replication of the input pixel
Rem                        onto the corresponding output pixels (box filter);
Rem  IJL_TRIANGLE_FILTER - 3/4 * nearer pixel + 1/4 * further pixel in each
Rem                        dimension
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum upsampling_type
 Public Const IJL_BOX_FILTER = 0
  Public Const IJL_TRIANGLE_FILTER = 1
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem Name:        SAMPLING_STATE
Rem
Rem Purpose:     Stores current conditions of sampling. Only for upsampling
Rem              with triangle filter is used now.
Rem
Rem Fields:
Rem  top_row        - pointer to buffer with MCUs, that are located above than
Rem                   current row of MCUs;
Rem  cur_row        - pointer to buffer with current row of MCUs;
Rem  bottom_row     - pointer to buffer with MCUs, that are located below than
Rem                   current row of MCUs;
Rem  last_row       - pointer to bottom boundary of last row of MCUs
Rem  cur_row_number - number of row of MCUs, that is decoding;
Rem  user_interrupt - field to store jprops->interrupt, because of we prohibit
Rem                   interrupts while top row of MCUs is upsampling.
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type SAMPLING_STATE
  top_row        As Long
  cur_row        As Long
  bottom_row     As Long
  last_row       As Long
  cur_row_number As Long
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        PROCESSOR_TYPE
Rem
Rem  Purpose:     Possible types of processors.
Rem               Note that the enums are defined in ascending order
Rem               depending upon their various IA32 instruction support.
Rem
Rem  Fields:
Rem
Rem    IJL_OTHER_PROC
Rem      Does not support the CPUID instruction and
Rem      assumes no Pentium(R) processor instructions.
Rem
Rem    IJL_PENTIUM_PROC
Rem      Corresponds to an Intel(R) Pentium(R) processor
Rem      (or a 100% compatible) that supports the
Rem      Pentium(R) processor instructions.
Rem
Rem    IJL_PENTIUM_PRO_PROC
Rem      Corresponds to an Intel(R) Pentium(R) Pro processor
Rem      (or a 100% compatible) that supports the
Rem      Pentium(R) Pro processor instructions.
Rem
Rem    IJL_PENTIUM_PROC_MMX_TECH
Rem      Corresponds to an Intel(R) Pentium(R) processor
Rem      with MMX(TM) technology (or a 100% compatible)
Rem      that supports the MMX(TM) instructions.
Rem
Rem    IJL_PENTIUM_II_PROC
Rem      Corresponds to an Intel(R) Pentium(R) II processor
Rem      (or a 100% compatible) that supports both the
Rem      Pentium(R) Pro processor instructions and the
Rem      MMX(TM) instructions.
Rem
Rem    IJL_PENTIUM_III_PROC
Rem      Corresponds to an Intel(R) Pentium(R) III processor
Rem
Rem  Any additional processor types that support a superset
Rem  of both the Pentium(R) Pro processor instructions and the
Rem  MMX(TM) instructions should be given an enum value greater
Rem  than IJL_PENTIUM_III_PROC.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum PROCESSOR_TYPE
 Public Const IJL_OTHER_PROC = 0
  Public Const IJL_PENTIUM_PROC = 1
  Public Const IJL_PENTIUM_PRO_PROC = 2
  Public Const IJL_PENTIUM_PROC_MMX_TECH = 3
  Public Const IJL_PENTIUM_II_PROC = 4
 Public Const IJL_PENTIUM_III_PROC = 5
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        ENTROPYSTRUCT
Rem
Rem  Purpose:     Stores the decoder state information necessary to "jump"
Rem               to a particular MCU row in a compressed entropy stream.
Rem
Rem  Context:     Used to persist the decoder state within Decode_Scan when
Rem               decoding using ROIs.
Rem
Rem  Fields:
Rem    offset              Offset (in bytes) into the entropy stream
Rem                        from the beginning.
Rem    dcval1              DC val at the beginning of the MCU row
Rem                        for component 1.
Rem    dcval2              DC val at the beginning of the MCU row
Rem                        for component 2.
Rem    dcval3              DC val at the beginning of the MCU row
Rem                        for component 3.
Rem    dcval4              DC val at the beginning of the MCU row
Rem                        for component 4.
Rem    bit_buffer_64       64-bit Huffman bit buffer.  Stores current
Rem                        bit buffer at the start of a MCU row.
Rem                        Also used as a 32-bit buffer on 32-bit
Rem                        architectures.
Rem    bitbuf_bits_valid   Number of valid bits in the above bit buffer.
Rem    unread_marker       Have any markers been decoded but not
Rem                        processed at the beginning of a MCU row?
Rem                        This entry holds the unprocessed marker, or
Rem                        0 if none.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type ENTROPYSTRUCT
  offset               As Long
  dcval1               As Long
  dcval2               As Long
  dcval3               As Long
  dcval4               As Long
  ' IJL use 8 byte pack structures
  pad0                 As Byte
  pad1                 As Byte
  pad2                 As Byte
  pad3                 As Byte
  bit_buffer_64        As Long
  bit_buffer_64_part_2 As Long
  bitbuf_bits_valid    As Long
  unread_marker        As Byte
  ' IJL use 8 byte pack structures
  pad4                 As Byte
  pad5                 As Byte
  pad6                 As Byte
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        STATE
Rem
Rem  Purpose:     Stores the active state of the IJL.
Rem
Rem  Context:     Used by all low-level routines to store pseudo-global or
Rem               state variables.
Rem
Rem  Fields:
Rem    bit_buffer_64           64-bit bitbuffer utilized by Huffman
Rem                            encoder/decoder algorithms utilizing routines
Rem                            designed for MMX(TM) technology.
Rem    bit_buffer_32           32-bit bitbuffer for all other Huffman
Rem                            encoder/decoder algorithms.
Rem    bitbuf_bits_valid       Number of bits in the above two fields that
Rem                            are valid.
Rem
Rem    cur_entropy_ptr         Current position (absolute address) in
Rem                            the entropy buffer.
Rem    start_entropy_ptr       Starting position (absolute address) of
Rem                            the entropy buffer.
Rem    end_entropy_ptr         Ending position (absolute address) of
Rem                            the entropy buffer.
Rem    entropy_bytes_processed Number of bytes actually processed
Rem                            (passed over) in the entropy buffer.
Rem    entropy_buf_maxsize     Max size of the entropy buffer.
Rem    entropy_bytes_left      Number of bytes left in the entropy buffer.
Rem    Prog_EndOfBlock_Run     Progressive block run counter.
Rem
Rem    DIB_ptr                 Temporary offset into the input/output DIB.
Rem
Rem    unread_marker           If a marker has been read but not processed,
Rem                            stick it in this field.
Rem    processor_type          (0, 1, or 2) == current processor does not
Rem                            support MMX(TM) instructions.
Rem                           (3 or 4) == current processor does
Rem                            support MMX(TM) instructions.
Rem    cur_scan_comp           On which component of the scan are we working?
Rem    file                    Process file handle, or
Rem                            0x00000000 if no file is defined.
Rem    JPGBuffer               Entropy buffer (~4K).
Rem
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type State
  bit_buffer_64                As Long
  bit_buffer_64_part_2         As Long
  bit_buffer_32                As Long
  bitbuf_bits_valid            As Long
  cur_entropy_ptr              As Long
  start_entropy_ptr            As Long
  end_entropy_ptr              As Long
  entropy_bytes_processed      As Long
  entropy_buf_maxsize          As Long
  entropy_bytes_left           As Long
  Prog_EndOfBlock_Run          As Long
  DIB_ptr                      As Long
  unread_marker                As Byte
  ' IJL use 8 byte pack structures
  pad0                         As Byte
  pad1                         As Byte
  pad2                         As Byte
  proc_type                    As Long 'PROCESSOR_TYPE
  cur_scan_comp                As Long
  file                         As Long
  JPGBuffer(0 To JBUFSIZE - 1) As Byte
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        FAST_MCU_PROCESSING_TYPE
Rem
Rem  Purpose:     Advanced Control Option.  Do NOT modify.
Rem               WARNING:  Used for internal reference only.
Rem
Rem  Fields:
Rem
Rem    IJL_(sampling)_(JPEG color space)_(sampling)_(DIB color space)
Rem      Decode is read left to right w/ upsampling.
Rem      Encode is read right to left w/ subsampling.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum FAST_MCU_PROCESSING_TYPE
 Public Const IJL_NO_CC_OR_US = 0

 Public Const IJL_111_YCBCR_111_RGB = 1
  Public Const IJL_111_YCBCR_111_BGR = 2

  Public Const IJL_411_YCBCR_111_RGB = 3
  Public Const IJL_411_YCBCR_111_BGR = 4

  Public Const IJL_422_YCBCR_111_RGB = 5
  Public Const IJL_422_YCBCR_111_BGR = 6

  Public Const IJL_111_YCBCR_1111_RGBA_FPX = 7
  Public Const IJL_411_YCBCR_1111_RGBA_FPX = 8
  Public Const IJL_422_YCBCR_1111_RGBA_FPX = 9

  Public Const IJL_1111_YCBCRA_FPX_1111_RGBA_FPX = 10
  Public Const IJL_4114_YCBCRA_FPX_1111_RGBA_FPX = 11
  Public Const IJL_4224_YCBCRA_FPX_1111_RGBA_FPX = 12

  Public Const IJL_111_RGB_1111_RGBA_FPX = 13

  Public Const IJL_1111_RGBA_FPX_1111_RGBA_FPX = 14

  Public Const IJL_111_OTHER_111_OTHER = 15
  Public Const IJL_411_OTHER_111_OTHER = 16
  Public Const IJL_422_OTHER_111_OTHER = 17

  Public Const IJL_YCBYCR_YCBCR = 18
'End Enum


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        JPEG_PROPERTIES
Rem
Rem  Purpose:     Stores low-level and control information.  It is used by
Rem               both the encoder and decoder.  An advanced external user
Rem               may access this structure to expand the interface
Rem               capability.
Rem
Rem               See the Developer's Guide for an expanded description
Rem               of this structure and its use.
Rem
Rem  Context:     Used by all interface methods and most IJL routines.
Rem
Rem  Fields:
Rem
Rem    iotype              IN:     Specifies type of data operation
Rem                                (read/write/other) to be
Rem                                performed by IJL_Read or IJL_Write.
Rem    roi                 IN:     Rectangle-Of-Interest to read from, or
Rem                                write to, in pixels.
Rem    dcttype             IN:     DCT alogrithm to be used.
Rem    fast_processing     OUT:    Supported fast pre/post-processing path.
Rem                                This is set by the IJL.
Rem    interrupt           IN:     Signals an interrupt has been requested.
Rem
Rem    DIBBytes            IN:     Pointer to buffer of uncompressed data.
Rem    DIBWidth            IN:     Width of uncompressed data.
Rem    DIBHeight           IN:     Height of uncompressed data.
Rem    DIBPadBytes         IN:     Padding (in bytes) at end of each
Rem                                row in the uncompressed data.
Rem    DIBChannels         IN:     Number of components in the
Rem                                uncompressed data.
Rem    DIBColor            IN:     Color space of uncompressed data.
Rem    DIBSubsampling      IN:     Required to be IJL_NONE or IJL_422.
Rem    DIBLineBytes        OUT:    Number of bytes in an output DIB line
Rem                                including padding.
Rem
Rem    JPGFile             IN:     Pointer to file based JPEG.
Rem    JPGBytes            IN:     Pointer to buffer based JPEG.
Rem    JPGSizeBytes        IN:     Max buffer size. Used with JPGBytes.
Rem                      OUT:      Number of compressed bytes written.
Rem    JPGWidth            IN:     Width of JPEG image.
Rem                      OUT:      After reading (except READHEADER).
Rem    JPGHeight           IN:     Height of JPEG image.
Rem                      OUT:      After reading (except READHEADER).
Rem    JPGChannels         IN:     Number of components in JPEG image.
Rem                      OUT:      After reading (except READHEADER).
Rem    JPGColor            IN:     Color space of JPEG image.
Rem    JPGSubsampling      IN:     Subsampling of JPEG image.
Rem                       OUT:     After reading (except READHEADER).
Rem    JPGThumbWidth       OUT:    JFIF embedded thumbnail width [0-255].
Rem    JPGThumbHeight      OUT:    JFIF embedded thumbnail height [0-255].
Rem
Rem    cconversion_reqd    OUT:    If color conversion done on decode, TRUE.
Rem    upsampling_reqd     OUT:    If upsampling done on decode, TRUE.
Rem    jquality            IN:     [0-100] where highest quality is 100.
Rem    jinterleaveType     IN/OUT: 0 => MCU interleaved file, and
Rem                                1 => 1 scan per component.
Rem    numxMCUs            OUT:    Number of MCUs in the x direction.
Rem    numyMCUs            OUT:    Number of MCUs in the y direction.
Rem
Rem    nqtables            IN/OUT: Number of quantization tables.
Rem    maxquantindex       IN/OUT: Maximum index of quantization tables.
Rem    nhuffActables       IN/OUT: Number of AC Huffman tables.
Rem    nhuffDctables       IN/OUT: Number of DC Huffman tables.
Rem    maxhuffindex        IN/OUT: Maximum index of Huffman tables.
Rem    jFmtQuant           IN/OUT: Formatted quantization table info.
Rem    jFmtAcHuffman       IN/OUT: Formatted AC Huffman table info.
Rem    jFmtDcHuffman       IN/OUT: Formatted DC Huffman table info.
Rem
Rem    jEncFmtQuant        IN/OUT: Pointer to one of the above, or
Rem                                to externally persisted table.
Rem    jEncFmtAcHuffman    IN/OUT: Pointer to one of the above, or
Rem                                to externally persisted table.
Rem    jEncFmtDcHuffman    IN/OUT: Pointer to one of the above, or
Rem                                to externally persisted table.
Rem
Rem    use_default_qtables IN:     Set to default quantization tables.
Rem                                Clear to supply your own.
Rem    use_default_htables IN:     Set to default Huffman tables.
Rem                                Clear to supply your own.
Rem    rawquanttables      IN:     Up to 4 sets of quantization tables.
Rem    rawhufftables       IN:     Alternating pairs (DC/AC) of up to 4
Rem                                sets of raw Huffman tables.
Rem    HuffIdentifierAC    IN:     Indicates what channel the user-
Rem                                supplied Huffman AC tables apply to.
Rem    HuffIdentifierDC    IN:     Indicates what channel the user-
Rem                                supplied Huffman DC tables apply to.
Rem
Rem    jframe              OUT:    Structure with frame-specific info.
Rem    needframe           OUT:    TRUE when a frame has been detected.
Rem
Rem    jscan               Persistence for current scan pointer when
Rem                        interrupted.
Rem
Rem    state               OUT:    Contains info on the state of the IJL.
Rem    SawAdobeMarker      OUT:    Decoder saw an APP14 marker somewhere.
Rem    AdobeXform          OUT:    If SawAdobeMarker TRUE, this indicates
Rem                                the JPEG color space given by that marker.
Rem
Rem    rowoffsets          Persistence for the decoder MCU row origins
Rem                        when decoding by ROI.  Offsets (in bytes
Rem                        from the beginning of the entropy data)
Rem                        to the start of each of the decoded rows.
Rem                        Fill the offsets with -1 if they have not
Rem                        been initalized and NULL could be the
Rem                        offset to the first row.
Rem
Rem    MCUBuf              OUT:    Quadword aligned internal buffer.
Rem                                Big enough for the largest MCU
Rem                                (10 blocks) with extra room for
Rem                                additional operations.
Rem    tMCUBuf             OUT:    Version of above, without alignment.
Rem
Rem    processor_type      OUT:    Determines type of processor found
Rem                                during initialization.
Rem
Rem    ignoreDCTs          IN:     Assert to bypass DCTs when processing
Rem                                data.  Required for conformance
Rem                                testing.
Rem
Rem    progressive_found   OUT:    1 when progressive image detected.
Rem    coef_buffer         IN:     Pointer to a larger buffer containing
Rem                                frequency coefficients when they
Rem                                cannot be decoded dynamically
Rem                                (i.e., as in progressive decoding).
Rem
Rem    upsampling_type     IN:     Type of sampling:
Rem                              IJL_BOX_FILTER or IJL_TRIANGLE_FILTER.
Rem    SAMPLING_STATE*    OUT:     pointer to structure, describing current
Rem                              condition of upsampling
Rem
Rem    AdobeVersion       OUT      version field, if Adobe APP14 marker detected
Rem    AdobeFlags0        OUT      flags0 field, if Adobe APP14 marker detected
Rem    AdobeFlags1        OUT      flags1 field, if Adobe APP14 marker detected
Rem
Rem    jfif_app0_detected OUT:     1 - if JFIF APP0 marker detected,
Rem                                0 - if not
Rem    jfif_app0_version  IN/OUT   The JFIF file version
Rem    jfif_app0_units    IN/OUT   units for the X and Y densities
Rem                                0 - no units, X and Y specify
Rem                                    the pixel aspect ratio
Rem                                1 - X and Y are dots per inch
Rem                                2 - X and Y are dots per cm
Rem    jfif_app0_Xdensity IN/OUT   horizontal pixel density
Rem    jfif_app0_Ydensity IN/OUT   vertical pixel density
Rem
Rem    jpeg_comment       IN       pointer to JPEG comments
Rem    jpeg_comment_size  IN/OUT   size of JPEG comments, in bytes
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type JPEG_PROPERTIES
  ' Compression/Decompression control.
  iotype                   As Long 'IJLIOTYPE                ' default = IJL_SETUP
  roi                      As Long 'IJL_RECT                 ' default = 0
  dct_type                 As Long 'DCTTYPE                  ' default = IJL_AAN
  fast_processing          As Long 'FAST_MCU_PROCESSING_TYPE ' default = IJL_NO_CC_OR_US
  interrupt                As Long                     ' default = FALSE
    
  ' DIB specific I/O data specifiers.
  DIBBytes                 As Long                     ' default = NULL
  DIBWidth                 As Long                     ' default = 0
  DIBHeight                As Long                     ' default = 0
  DIBPadBytes              As Long                     ' default = 0
  DIBChannels              As Long                     ' default = 3
  DIBColor                 As Long 'IJL_COLOR                ' default = IJL_BGR
  DIBSubsampling           As Long 'IJL_DIBSUBSAMPLING       ' default = IJL_NONE
  DIBLineBytes             As Long                     ' default = 0
    
  ' JPEG specific I/O data specifiers.
  JPGFile                  As Long                     ' default = NULL
  JPGBytes                 As Long                     ' default = NULL
  JPGSizeBytes             As Long                     ' default = 0
  JPGWidth                 As Long                     ' default = 0
  JPGHeight                As Long                     ' default = 0
  JPGChannels              As Long                     ' default = 3
  JPGColor                 As Long 'IJL_COLOR                ' default = IJL_YCBCR
  JPGSubsampling           As Long 'IJL_JPGSUBSAMPLING       ' default = IJL_411
  JPGThumbWidth            As Long                     ' default = 0
  JPGThumbHeight           As Long                     ' default = 0
    
  ' JPEG conversion properties.
  cconversion_reqd         As Long                     ' default = TRUE
  upsampling_reqd          As Long                     ' default = TRUE
  jquality                 As Long                     ' default = 75
  
  
    '// Low-level properties - 20,000 bytes.  If the whole structure
  ' is written out then VB fails with an obscure error message
  ' "Too Many Local Variables" !
  '
  ' These all default if they are not otherwise specified so there
  ' is no trouble to just assign a sufficient buffer in memory:
  jpropsLL(0 To 19999) As Byte

  
  
  
  
  
  'ннннннннннннннннннннннннннннннннннннннннннннннннна
'аjinterleaveType          As Long                     ' default = 0
'а  numxMCUs                 As Long                     ' default = 0
'а  numyMCUs                 As Long                     ' default = 0
'а
'а  ' Tables.
'а  nqtables                 As Long
'а  maxquantindex            As Long
'а  nhuffActables            As Long
'а  nhuffDctables            As Long
'а  maxhuffindex             As Long
'а
'а  jFmtQuant(0 To 3)        As QUANT_TABLE
'а  jFmtAcHuffman(0 To 3)    As HUFFMAN_TABLE
'а  jFmtDcHuffman(0 To 3)    As HUFFMAN_TABLE
'а
'а  jEndFmtQuant(0 To 3)     As Long
'а  jEncFmtAcHuffman(0 To 3) As Long
'а  jEndFmtDcHuffman(0 To 3) As Long
'а
'а  ' Allow user-defined tables.
'а  use_default_qtables      As Long
'а  use_default_htables      As Long
'а
'а  rawquanttables(0 To 3)   As JPEGQuantTable
'а  rawhufftables(0 To 7)    As JPEGHuffTable
'а  HuffIdentifierAC(0 To 3) As Byte
'а  HuffIdentifierDC(0 To 3) As Byte
'а
'а  ' Frame specific members.
'а  jframe                   As FRAME
'а  needframe                As Long
'а
'а  ' SCAN persistent members.
'а  jscan                    As Long
'а
'а  ' IJL use 8 byte pack structures
'а  pad0                     As Byte
'а  pad1                     As Byte
'а  pad2                     As Byte
'а  pad3                     As Byte
'а
'а  ' State members.
'а  state_field              As STATE
'а  SawAdobeMarker           As Long
'а  AdobeXform               As Long
'а
'а  ' ROI decoder members.
'а  rowoffsets               As Long
'а
'а  ' Intermediate buffers.
'а  MCUBuf                   As Long
'а  tMCUBuf(0 To 1439)       As Byte
'а
'а  ' Processor detected.
'а  processortype            As Long 'PROCESSOR_TYPE
'а
'а  ' Test specific members.
'а  ignoreDCTs               As Long
'а
'а  ' Progressive mode members.
'а  progressive_found        As Long
'а  coef_buffer              As Long
'а
'а  ' Upsampling mode members.
'а  upsampling_type          As Long 'upsampling_type
'а  sampling_state_ptr       As Long
'а
'а  ' Adobe APP14 segment variables
'а  AdobeVersion             As Integer
'а  AdobeFlags0              As Integer
'а  AdobeFlags1              As Integer
'а
'а  ' JFIF APP0 segment variables
'а  jfif_app0_detected       As Long
'а  jfif_app0_version        As Integer
'а  jfif_app0_units          As Byte
'а  jfif_app0_Xdensity       As Integer
'а  jfif_app0_Ydensity       As Integer
'а
'а  ' comments related fields
'а  jpeg_comment             As Long
'а  jpeg_comment_size        As Integer
'а
'ннннннннннннннннннннннннннннннннннннннннннннннннна

End Type


Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        JPEG_CORE_PROPERTIES
Rem
Rem  Purpose:     This is the primary data structure between the IJL and
Rem               the external user.  It stores JPEG state information
Rem               and controls the IJL.  It is user-modifiable.
Rem
Rem               See the Developer's Guide for details on appropriate usage.
Rem
Rem  Context:     Used by all low-level IJL routines to store
Rem               pseudo-global information.
Rem
Rem  Fields:
Rem
Rem    UseJPEGPROPERTIES   Set this flag != 0 if you wish to override
Rem                        the JPEG_CORE_PROPERTIES "IN" parameters with
Rem                        the JPEG_PROPERTIES parameters.
Rem
Rem    DIBBytes            IN:     Pointer to buffer of uncompressed data.
Rem    DIBWidth            IN:     Width of uncompressed data.
Rem    DIBHeight           IN:     Height of uncompressed data.
Rem    DIBPadBytes         IN:     Padding (in bytes) at end of each
Rem                                row in the uncompressed data.
Rem    DIBChannels         IN:     Number of components in the
Rem                                uncompressed data.
Rem    DIBColor            IN:     Color space of uncompressed data.
Rem    DIBSubsampling      IN:     Required to be IJL_NONE or IJL_422.
Rem
Rem    JPGFile             IN:     Pointer to file based JPEG.
Rem    JPGBytes            IN:     Pointer to buffer based JPEG.
Rem    JPGSizeBytes        IN:     Max buffer size. Used with JPGBytes.
Rem                        OUT:    Number of compressed bytes written.
Rem    JPGWidth            IN:     Width of JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGHeight           IN:     Height of JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGChannels         IN:     Number of components in JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGColor            IN:     Color space of JPEG image.
Rem    JPGSubsampling      IN:     Subsampling of JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGThumbWidth       OUT:    JFIF embedded thumbnail width [0-255].
Rem    JPGThumbHeight      OUT:    JFIF embedded thumbnail height [0-255].
Rem
Rem    cconversion_reqd    OUT:    If color conversion done on decode, TRUE.
Rem    upsampling_reqd     OUT:    If upsampling done on decode, TRUE.
Rem    jquality            IN:     [0-100] where highest quality is 100.
Rem
Rem    jprops              "Low-Level" IJL data structure.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Type JPEG_CORE_PROPERTIES
  UseJPEGPROPERTIES As Long               ' default = 0
    
  ' DIB specific I/O data specifiers.
  DIBBytes          As Long               ' default = NULL
  DIBWidth          As Long               ' default = 0
  DIBHeight         As Long               ' default = 0
  DIBPadBytes       As Long               ' default = 0
  DIBChannels       As Long               ' default = 3
  DIBColor          As Long 'IJL_COLOR          ' default = IJL_BGR
  DIBSubsampling    As Long 'IJL_DIBSUBSAMPLING ' default = IJL_NONE
    
  ' JPEG specific I/O data specifiers.
  JPGFile           As Long               ' default = NULL
  JPGBytes          As Long               ' default = NULL
  JPGSizeBytes      As Long               ' default = 0
  JPGWidth          As Long               ' default = 0
  JPGHeight         As Long               ' default = 0
  JPGChannels       As Long               ' default = 3
  JPGColor          As Long 'IJL_COLOR          ' default = IJL_YCBCR
  JPGSubsampling    As Long 'IJL_JPGSUBSAMPLING ' default = IJL_411
  JPGThumbWidth     As Long               ' default = 0
  JPGThumbHeight    As Long               ' default = 0
    
  ' JPEG conversion properties.
  cconversion_reqd  As Long               ' default = TRUE
  upsampling_reqd   As Long               ' default = TRUE
  jquality          As Long               ' default = 75
  
  ' IJL use 8 byte pack structures
  pad0              As Byte
  pad1              As Byte
  pad2              As Byte
  pad3              As Byte
  
  ' Low-level properties.
  jprops            As JPEG_PROPERTIES
  
End Type


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        IJLERR
Rem
Rem  Purpose:     Listing of possible "error" codes returned by the IJL.
Rem
Rem               See the Developer's Guide for details on appropriate usage.
Rem
Rem  Context:     Used for error checking.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

'Public Enum IJLERR
  ' The following "error" values indicate an "OK" condition.
 Public Const IJL_OK = 0
 Public Const IJL_INTERRUPT_OK = 1
 Public Const IJL_ROI_OK = 2
    
  ' The following "error" values indicate an error has occurred.
 Public Const IJL_EXCEPTION_DETECTED = -1
  Public Const IJL_INVLAID_ENCODER = -2
  Public Const IJL_UNSUPPORTED_SUBSAMPLING = -3
 Public Const IJL_UNSUPPORTED_BYTES_PER_PIXEL = -4
 Public Const IJL_MEMORY_ERROR = -5
 Public Const IJL_BAD_HUFFMAN_TABLE = -6
 Public Const IJL_BAD_QUANT_TABLE = -7
 Public Const IJL_INVALID_JPEG_PROPERTIES = -8
 Public Const IJL_ERR_FILECLOSE = -9
Public Const IJL_INVALID_FILENAME = -10
Public Const IJL_ERROR_EOF = -11
 Public Const IJL_PROG_NOT_SUPPORTED = -12
 Public Const IJL_ERR_NOT_JPEG = -13
 Public Const IJL_ERR_COMP = -14
 Public Const IJL_ERR_SOF = -15
 Public Const IJL_ERR_DNL = -16
 Public Const IJL_ERR_NO_HUF = -17
 Public Const IJL_ERR_NO_QUAN = -18
 Public Const IJL_ERR_NO_FRAME = -19
 Public Const IJL_ERR_MULT_FRAME = -20
 Public Const IJL_ERR_DATA = -21
Public Const IJL_ERR_NO_IMAGE = -22
 Public Const IJL_FILE_ERROR = -23
 Public Const IJL_INTERNAL_ERROR = -24
  Public Const IJL_BAD_RST_MARKER = -25
Public Const IJL_THUMBNAIL_DIB_TOO_SMALL = -26
 Public Const IJL_THUMBNAIL_DIB_WRONG_COLOR = -27
 Public Const IJL_BUFFER_TOO_SMALL = -28
 Public Const IJL_UNSUPPORTED_FRAME = -29
 Public Const IJL_ERR_COM_BUFFER = -30
 Public Const IJL_RESERVED = -99
'End Enum


Public Declare Function ijlInit Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES) As Long 'IJLERR
Public Declare Function ijlFree Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES) As Long 'IJLERR
Public Declare Function ijlRead Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES, ByVal iotype As Long) As Long  'IJLERR
Public Declare Function ijlWrite Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES, ByVal iotype As Long) As Long 'IJLERR
Public Declare Function ijlGetLibVersion Lib "ijl15.dll" () As Long 'pointer to IJLibVersion...
Public Declare Function ijlErrorStr Lib "ijl15.dll" (Code As Long) As Long  'pointer to C-style string


Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem  Name:        IJL_DIB_PAD_BYTES
Rem
Rem  Purpose:     Calculate number of bytes to pad DIB line.
Rem
Rem//////////////////////////////////////////////////////////////////////////
Rem

Public Function IJL_DIB_PAD_BYTES(ByVal width As Long, ByVal nChannels As Long) As Long
Dim IJL_DIB_ALIGN As Long
Dim IJL_DIB_UWIDTH As Long
Dim IJL_DIB_AWIDTH As Long

  IJL_DIB_ALIGN = 3
  IJL_DIB_UWIDTH = width * nChannels
  IJL_DIB_AWIDTH = (IJL_DIB_UWIDTH + IJL_DIB_ALIGN) And (Not (IJL_DIB_ALIGN))
  
  IJL_DIB_PAD_BYTES = IJL_DIB_AWIDTH - IJL_DIB_UWIDTH

End Function