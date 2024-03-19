#pragma once  

#include <fstream>
#include <iostream>
#include <string>
#include <sstream>
#include <vector>
#include <libxl.h>  
#include <filesystem>
#include <windows.h>
#include <CommCtrl.h>

using namespace libxl; 

#ifndef TOOLCLASS_H
#define TOOLCLASS_H

class XLSXUtils {
public:
	static BOOL GetAllFiles(LPWSTR path, std::vector<std::string> &files, std::vector<std::string>& XMLfiles, std::vector<std::string>& PDFfiles);
	static void DoContain(std::string ID, std::string name, std::vector<std::string>& files, std::vector<std::string>& XMLfiles, HWND hWndProgres);
	static void Export(std::string outputPath, std::vector<std::string>& files, std::vector<std::string>& XMLfiles, HWND hWndProgress);
};

#endif


