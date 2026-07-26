#pragma once

/* Windows/MSVC configuration for the embedded libxls 1.6.3 build. */
#define PACKAGE "libxls"
#define PACKAGE_NAME "libxls"
#define PACKAGE_TARNAME "libxls"
#define PACKAGE_VERSION "1.6.3"
#define PACKAGE_STRING "libxls 1.6.3"
#define PACKAGE_BUGREPORT ""
#define PACKAGE_URL "https://github.com/libxls/libxls"
#define VERSION "1.6.3"

#define STDC_HEADERS 1
#define HAVE_INTTYPES_H 1
#define HAVE_MEMORY_H 1
#define HAVE_STDINT_H 1
#define HAVE_STDLIB_H 1
#define HAVE_STRING_H 1
#define HAVE_SYS_STAT_H 1
#define HAVE_SYS_TYPES_H 1

/* Windows is little-endian. Do not define WORDS_BIGENDIAN. */

#ifdef _MSC_VER
#ifndef __cplusplus
#ifndef inline
#define inline __inline
#endif
#endif
#ifndef strdup
#define strdup _strdup
#endif
#ifndef strcasecmp
#define strcasecmp _stricmp
#endif
#ifndef strncasecmp
#define strncasecmp _strnicmp
#endif
#endif
