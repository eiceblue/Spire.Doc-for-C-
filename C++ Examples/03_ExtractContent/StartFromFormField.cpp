#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextInputField.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"StartFromFormField.docx";

	//Create the source document
	Document* sourceDocument = new Document();

	//Load the source document from disk.
	sourceDocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	Document* destinationDoc = new Document();

	//Add a section
	Section* section = destinationDoc->AddSection();

	//Define a variables
	int index = 0;

	//Traverse FormFields
	for (int i = 0; i < sourceDocument->GetSections()->GetItem(0)->GetBody()->GetFormFields()->GetCount(); i++)
	{
		FormField* field = sourceDocument->GetSections()->GetItem(0)->GetBody()->GetFormFields()->GetItem(i);
		//Find FieldFormTextInput type field
		if (field->GetType() == FieldType::FieldFormTextInput)
		{
			//Get the paragraph
			Paragraph* paragraph = field->GetOwnerParagraph();

			//Get the index
			index = sourceDocument->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->IndexOf(paragraph);
			break;
		}
	}

	//Extract the content
	for (int i = index; i < index + 3; i++)
	{
		//Clone the ChildObjects of source document
		DocumentObject* doobj = sourceDocument->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		section->GetBody()->GetChildObjects()->Add(doobj);
	}

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	sourceDocument->Close();
	destinationDoc->Close();
	delete sourceDocument;
	delete destinationDoc;
}
