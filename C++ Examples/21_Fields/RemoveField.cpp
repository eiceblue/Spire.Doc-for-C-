#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"IfFieldSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveField.docx";

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first field
	Field* field = document->GetFields()->GetItem(0);

	//Get the paragraph of the field
	Paragraph* par = field->GetOwnerParagraph();
	//Get the index of the  field
	int index = par->GetChildObjects()->IndexOf(field);
	//Remove if field via index
	par->GetChildObjects()->RemoveAt(index);

	//Save doc file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
