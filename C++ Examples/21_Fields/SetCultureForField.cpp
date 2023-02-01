#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetCultureForField.docx";

	//Create a document
	Document* document = new Document();

	//Create a section
	Section* section = document->AddSection();

	//Add paragraph
	Paragraph* paragraph = section->AddParagraph();

	//Add textRnage for paragraph
	paragraph->AppendText(L"Add Date Field: ");

	//Add date field1
	Field* field1 = dynamic_cast<Field*>(paragraph->AppendField(L"Date1", FieldType::FieldDate));
	wstring codeString = L"DATE  \\@";
	codeString += L"\"yyyy\\MM\\dd\"";
	field1->SetCode(codeString.c_str());

	//Add new paragraph
	Paragraph* newParagraph = section->AddParagraph();

	//Add textRnage for paragraph
	newParagraph->AppendText(L"Add Date Field with setting French Culture: ");

	//Add date field2
	Field* field2 = newParagraph->AppendField(L"\"\\@\"dd MMMM yyyy", FieldType::FieldDate);

	//Setting Field with setting French Culture
	field2->GetCharacterFormat()->SetLocaleIdASCII(1036);

	//Update fields
	document->SetIsUpdateFields(true);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
	delete field2;
}
