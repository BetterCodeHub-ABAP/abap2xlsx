INTERFACE zif_excel_row
  PUBLIC .
  METHODS get_collapsed
    IMPORTING
      !io_worksheet      TYPE REF TO zcl_excel_worksheet OPTIONAL
    RETURNING
      VALUE(r_collapsed) TYPE boolean .
  METHODS get_outline_level
    IMPORTING
      !io_worksheet          TYPE REF TO zcl_excel_worksheet OPTIONAL
    RETURNING
      VALUE(r_outline_level) TYPE int4 .
  METHODS get_row_height
    RETURNING
      VALUE(r_row_height) TYPE float .
  METHODS get_row_index
    RETURNING
      VALUE(r_row_index) TYPE int4 .
  METHODS get_visible
    IMPORTING
      !io_worksheet    TYPE REF TO zcl_excel_worksheet OPTIONAL
    RETURNING
      VALUE(r_visible) TYPE boolean .
  METHODS get_xf_index
    RETURNING
      VALUE(r_xf_index) TYPE int4 .
  METHODS set_collapsed
    IMPORTING
      !ip_collapsed TYPE boolean .
  METHODS set_outline_level
    IMPORTING
      !ip_outline_level TYPE int4
    RAISING
      zcx_excel .
  METHODS set_row_height
    IMPORTING
      !ip_row_height TYPE simple
    RAISING
      zcx_excel .
  METHODS set_row_index
    IMPORTING
      !ip_index TYPE int4 .
  METHODS set_visible
    IMPORTING
      !ip_visible TYPE boolean .
  METHODS set_xf_index
    IMPORTING
      !ip_xf_index TYPE int4 .
ENDINTERFACE.
