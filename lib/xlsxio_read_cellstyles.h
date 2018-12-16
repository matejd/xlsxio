#ifndef INCLUDED_XLSXIO_READ_CELLSTYLES_H
#define INCLUDED_XLSXIO_READ_CELLSTYLES_H

#include <stdint.h>
#include <expat.h>

#ifdef __cplusplus
extern "C" {
#endif

struct cell_styles_callback_data {
  XML_Parser xmlparser;
  int* cell_style_formats;
  int num_cell_style_formats;
};

#ifdef ASSUME_NO_NAMESPACE
#define XML_Char_icmp_ins XML_Char_icmp
#else
int XML_Char_icmp_ins (const XML_Char* value, const XML_Char* name);
#endif

const XML_Char* get_expat_attr_by_name (const XML_Char** attrs, const XML_Char* name);
void cell_styles_callback_data_initialize (struct cell_styles_callback_data* data, int* cell_style_formats);
void cell_styles_callback_find_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void cell_styles_callback_xf_handler (void* callbackdata, const XML_Char* name, const XML_Char** atts);

#ifdef __cplusplus
}
#endif

#endif
