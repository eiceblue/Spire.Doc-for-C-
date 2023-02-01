#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveReadOnlyRestriction.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveReadOnlyRestriction.docx";

	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	//Remove ReadOnly Restriction.
	doc->Protect(ProtectionType::NoProtection);
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}
