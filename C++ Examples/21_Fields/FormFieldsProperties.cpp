#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FillFormField.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FormFieldsProperties.docx";

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = document->GetSections()->GetItem(0);

	//Get FormField by index
	FormField* formField = section->GetBody()->GetFormFields()->GetItem(1);

	if (formField->GetType() == FieldType::FieldFormTextInput)
	{
		wstring formFieldName = formField->GetName();
		wstring temp = L"My name is " + formFieldName;
		formField->SetText(temp.c_str());
		formField->GetCharacterFormat()->SetTextColor(Color::GetRed());
		formField->GetCharacterFormat()->SetItalic(true);
	}

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
