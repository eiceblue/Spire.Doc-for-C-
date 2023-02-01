#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Fields.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ConvertFieldToText.docx";

	//Load word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get all fields in document
	FieldCollection* fields = document->GetFields();
	int count = fields->GetCount();

	for (int i = 0; i < count; i++)
	{
		Field* field = fields->GetItem(0);
		wstring s = field->GetFieldText();
		int index = field->GetOwnerParagraph()->GetChildObjects()->IndexOf(field);
		TextRange* textRange = new TextRange(document);
		textRange->SetText(s.c_str());
		textRange->GetCharacterFormat()->SetFontSize(24.0f);

		field->GetOwnerParagraph()->GetChildObjects()->Insert(index, textRange);
		field->GetOwnerParagraph()->GetChildObjects()->Remove(field);

	}

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}