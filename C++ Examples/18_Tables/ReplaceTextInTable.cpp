#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReplaceTextInTable.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceTextInTable.docx";

	//Load Word from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Get the first table in the section
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	//Define a regular expression to match the {} with its content
	Regex* regex = new Regex(L"{[^\\}]+\\}");

	//Replace the text of table with regex
	table->Replace(regex, L"E-iceblue");

	//Replace old text with new text in table
	table->Replace(L"Beijing", L"Component", false, true);

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}
