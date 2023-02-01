#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SetEditableRange.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetEditableRange.docx";

	//Create a new document
	Document* document = new Document();
	//Load file from disk
	document->LoadFromFile(inputFile.c_str());
	//Protect whole document
	document->Protect(ProtectionType::AllowOnlyReading, L"password");
	//Create tags for permission start and end
	PermissionStart* start = new PermissionStart(document, L"testID");
	PermissionEnd* end = new PermissionEnd(document, L"testID");
	//Add the start and end tags to allow the first paragraph to be edited.
	document->GetSections()->GetItem(0)->GetParagraphs()->GetItem(0)->GetChildObjects()->Insert(0, start);
	document->GetSections()->GetItem(0)->GetParagraphs()->GetItem(0)->GetChildObjects()->Add(end);
	//Save the document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}