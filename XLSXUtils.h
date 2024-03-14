#pragma once  

#include <libxl.h>  

using namespace libxl;  

Book* book = 0;

void Initialize() {
	book = xlCreateBook();
}
