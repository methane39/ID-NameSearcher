#include "XLSXUtils.h"
namespace fs = std::filesystem;

std::vector<std::string> DelXMLfiles;
std::vector<std::string> Delfiles;

BOOL ends_with(const std::string& str, const std::string& suffix) {
	if (suffix.length() > str.length()) {
		return false;
	}
	return std::equal(suffix.rbegin(), suffix.rend(), str.rbegin());
}

BOOL XLSXUtils::GetAllFiles(LPWSTR path, std::vector<std::string>& files, std::vector<std::string>& XMLfiles, std::vector<std::string>& PDFfiles) {


	std::wstring wstr(path);
	std::string str(wstr.begin(), wstr.end());

	for (const auto& entry : fs::recursive_directory_iterator(path)) {
		std::string file_path = entry.path().string();
		if (ends_with(file_path, ".xls")) {
			files.push_back(file_path);
		}
		else if (ends_with(file_path, ".xlsx")) {
			XMLfiles.push_back(file_path);
		}
		else if (ends_with(file_path, ".pdf")) {
			PDFfiles.push_back(file_path);
		}
	}
	return 1;
}

BOOL XMLDoHave(std::string name, std::string ID, std::string filepath) {

	Book* book = xlCreateXMLBook();
	book->load(filepath.c_str());
	Sheet* sheet = book->getSheet(0);

	for (int i = 0; i < sheet->lastRow() + 1; i++) {
		for (int j = 0; j < sheet->lastCol() + 1; j++) {
			if (const char* c_str = sheet->readStr(i, j)) {
				std::string gstr = std::string(c_str);
				auto it1 = std::search(gstr.begin(), gstr.end(), name.begin(), name.end());
				auto it2 = std::search(gstr.begin(), gstr.end(), ID.begin(), ID.end());
				if (it1 != gstr.end() || it2 != gstr.end()) {
					//delete sheet;
					//delete book;
					book->release();
					return 1;
				}
			}
		}
	}
	//delete sheet;
	//delete book;
	book->release();
	return 0;
}

BOOL DoHave(std::string name, std::string ID, std::string filepath) {

	Book* book = xlCreateBook();
	book->load(filepath.c_str());
	Sheet* sheet = book->getSheet(0);

	for (int i = 0; i < sheet->lastRow() + 1; i++) {
		for (int j = 0; j < sheet->lastCol() + 1; j++) {
			if (const char* c_str = sheet->readStr(i, j)) {
				std::string gstr = std::string(c_str);
				auto it1 = std::search(gstr.begin(), gstr.end(), name.begin(), name.end());
				auto it2 = std::search(gstr.begin(), gstr.end(), ID.begin(), ID.end());
				if (it1 != gstr.end() || it2 != gstr.end()) {
					//delete sheet;
					//delete book;
					book->release();
					return 1;
				}
			}
		}
	}
	//delete sheet;
	//delete book;
	book->release();
	return 0;
}

void XLSCheck(std::string name, std::string ID, std::vector<std::string>& files, HWND hWndProgress) {
	for each (std::string filepath in files)
	{
		if (!DoHave(name, ID, filepath)) {
			Delfiles.push_back(filepath);
		}
	}
	SendMessage(hWndProgress, PBM_SETPOS, 15, 0);
	for each (std::string delpath in Delfiles)
	{
		auto it = std::find(files.begin(), files.end(), delpath);
		files.erase(it);
	}
	SendMessage(hWndProgress, PBM_SETPOS, 25, 0);
}

void XLSXCheck(std::string name, std::string ID, std::vector<std::string>& XMLfiles, HWND hWndProgress) {
	for each (std::string filepath in XMLfiles)
	{
		if (!XMLDoHave(name, ID, filepath)) {
			DelXMLfiles.push_back(filepath);
		}	
	}
	SendMessage(hWndProgress, PBM_SETPOS, 55, 0);
	for each (std::string delpath in DelXMLfiles)
	{
		auto it = std::find(XMLfiles.begin(), XMLfiles.end(), delpath);
		XMLfiles.erase(it);
	}
	SendMessage(hWndProgress, PBM_SETPOS, 70, 0);
}

void PDFCheck(std::string name, std::string ID, std::vector<std::string>& PDFfiles) {

}

void XLSXUtils::DoContain(std::string ID, std::string name, std::vector<std::string>& files, std::vector<std::string>& XMLfiles, HWND hWndProgress) {
	XLSCheck(name, ID, files, hWndProgress);
	SendMessage(hWndProgress, PBM_SETPOS, 45, 0);
	XLSXCheck(name, ID, XMLfiles, hWndProgress);
	SendMessage(hWndProgress, PBM_SETPOS, 85, 0);

}

void XLSXUtils::Export(std::string outputPath, std::vector<std::string>& files, std::vector<std::string>& XMLfiles, HWND hWndProgress) {
	std::string folder_path = outputPath;
	CreateDirectory(folder_path.c_str(), NULL);
	SendMessage(hWndProgress, PBM_SETPOS, 90, 0);

	for each (const std::string file_path in files)
	{
		CopyFile(file_path.c_str(), (folder_path + "\\" + file_path).c_str(), FALSE);

		std::ofstream out(folder_path + "\\list.txt", std::ios::app);
		out << file_path << std::endl;
	}
	SendMessage(hWndProgress, PBM_SETPOS, 95, 0);

	for each (const std::string file_path in XMLfiles)
	{
		CopyFile(file_path.c_str(), (folder_path + "\\" + file_path).c_str(), FALSE);

		std::ofstream out(folder_path + "\\list.txt", std::ios::app);
		out << file_path << std::endl;
	}
	SendMessage(hWndProgress, PBM_SETPOS, 100, 0);
}