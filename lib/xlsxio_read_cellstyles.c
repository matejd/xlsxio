#include "xlsxio_private.h"
#include "xlsxio_read_cellstyles.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

void cell_styles_callback_data_initialize (struct cell_styles_callback_data* data, int* cell_style_formats)
{
  data->xmlparser = NULL;
  data->cell_style_formats = cell_style_formats;
  data->num_cell_style_formats = 0;
}

void cell_styles_callback_find_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct cell_styles_callback_data* data = (struct cell_styles_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("cellXfs")) == 0) {
    XML_SetElementHandler(data->xmlparser, cell_styles_callback_xf_handler, NULL);
  }
}

void cell_styles_callback_xf_handler (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct cell_styles_callback_data* data = (struct cell_styles_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("xf")) == 0) {
    const XML_Char* value = get_expat_attr_by_name(atts, X("numFmtId"));
    if (value != NULL) {
      sscanf(value, "%d", &data->cell_style_formats[data->num_cell_style_formats++]);
    }
  } else {
    XML_SetElementHandler(data->xmlparser, cell_styles_callback_find_start, NULL);
  }
}
