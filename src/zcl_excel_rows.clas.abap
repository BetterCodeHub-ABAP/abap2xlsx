*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_ROWS DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class ZCL_EXCEL_ROWS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_ROWS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  public section.

    methods ADD
      importing
        !IO_ROW type ref to ZCL_EXCEL_ROW .
    methods CLEAR .
    methods CONSTRUCTOR .
    methods GET
      importing
        !IP_INDEX     type I
      returning
        value(EO_ROW) type ref to ZCL_EXCEL_ROW .
    methods GET_ITERATOR
      returning
        value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
    methods IS_EMPTY
      returning
        value(IS_EMPTY) type FLAG .
    methods REMOVE
      importing
        !IO_ROW type ref to ZCL_EXCEL_ROW .
    methods SIZE
      returning
        value(EP_SIZE) type I .
  protected section.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  private section.

    data ROWS type ref to CL_OBJECT_COLLECTION .
endclass.



class ZCL_EXCEL_ROWS implementation.


  method ADD.
    ROWS->ADD( IO_ROW ).
  endmethod.                    "ADD


  method CLEAR.
    ROWS->CLEAR( ).
  endmethod.                    "CLEAR


  method CONSTRUCTOR.

    create object ROWS.

  endmethod.                    "CONSTRUCTOR


  method GET.
    EO_ROW ?= ROWS->IF_OBJECT_COLLECTION~GET( IP_INDEX ).
  endmethod.                    "GET


  method GET_ITERATOR.
    EO_ITERATOR ?= ROWS->IF_OBJECT_COLLECTION~GET_ITERATOR( ).
  endmethod.                    "GET_ITERATOR


  method IS_EMPTY.
    IS_EMPTY = ROWS->IF_OBJECT_COLLECTION~IS_EMPTY( ).
  endmethod.                    "IS_EMPTY


  method REMOVE.
    ROWS->REMOVE( IO_ROW ).
  endmethod.                    "REMOVE


  method SIZE.
    EP_SIZE = ROWS->IF_OBJECT_COLLECTION~SIZE( ).
  endmethod.                    "SIZE
endclass.
