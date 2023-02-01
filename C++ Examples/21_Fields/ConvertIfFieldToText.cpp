#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"IfFieldSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ConvertIfFieldToText.docx";

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get all fields in document
	FieldCollection* fields = document->GetFields();

	for (int i = 0; i < fields->GetCount(); i++)
	{
		Field* field = fields->GetItem(i);
		if (field->GetType() == FieldType::FieldIf)
		{
			TextRange* original = dynamic_cast<TextRange*>(field);
			//Get field text
			wstring text = field->GetFieldText();
			//Create a new textRange and set its format
			TextRange* textRange = new TextRange(document);
			textRange->SetText(text.c_str());
			textRange->GetCharacterFormat()->SetFontName(original->GetCharacterFormat()->GetFontName());
			textRange->GetCharacterFormat()->SetFontSize(original->GetCharacterFormat()->GetFontSize());

			Paragraph* par = field->GetOwnerParagraph();
			//Get the index of the if field
			int index = par->GetChildObjects()->IndexOf(field);
			//Remove if field via index
			par->GetChildObjects()->RemoveAt(index);
			//Insert field text at the position of if field
			par->GetChildObjects()->Insert(index, textRange);
		}
	}
	//Save doc file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
