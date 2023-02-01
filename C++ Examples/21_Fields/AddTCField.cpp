#include "pch.h"
using namespace Spire::Doc;

int main(){
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddTCField.docx";
	
	//Create Word document.
	Document* document = new Document();

	//Add a new section.
	Section* section = document->AddSection();

	//Add a new paragraph.
	Paragraph* paragraph = section->AddParagraph();

	//Add TC field in the paragraph
	Field* field = paragraph->AppendField(L"TC", FieldType::FieldTOCEntry);
	wstring codeString = L"TC ";
	codeString += L"\"Entry Text\"";
	codeString += L" \\f";
	codeString += L" t";
	field->SetCode(codeString.c_str());
	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

