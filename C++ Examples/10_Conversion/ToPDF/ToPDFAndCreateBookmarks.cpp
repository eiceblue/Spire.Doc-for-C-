#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"BookmarkTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToPDFAndCreateBookmarks.pdf";

	Document* document = new Document();
	//Load the document from disk
	document->LoadFromFile(inputFile.c_str());
	ToPdfParameterList* parames = new ToPdfParameterList();
	//Set CreateWordBookmarks to true
	parames->SetCreateWordBookmarks(true);
	//Create bookmarks using Headings
	parames->SetCreateWordBookmarksUsingHeadings(false);
	document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	document->Close();
	delete document;
}
