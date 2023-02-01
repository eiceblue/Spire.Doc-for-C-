#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReplaceTextByRegex.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceTextByRegex.docx";

	//create a document
	Document* doc = new Document();

	//Load the document from disk.
	doc->LoadFromFile(inputFile.c_str());

	//create a regex, match the text that starts with #
	Regex* regex = new Regex(L"\\#\\w+\\b");

	//replace the text by regex
	doc->Replace(regex, L"Spire.Doc");

	//save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
