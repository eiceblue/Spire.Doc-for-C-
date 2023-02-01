#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FillFormField.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetFormFieldsCollection.txt";

	wstring* sb = new wstring();

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = document->GetSections()->GetItem(0);

	FormFieldCollection* formFields = section->GetBody()->GetFormFields();

	sb->append(L"The first section has " + to_wstring(formFields->GetCount()) + L" form fields.");

	wofstream out;
	out.open(outputFile);
	out.flush();
	out << sb->c_str();
	out.close();

	document->Close();
	delete document;
	delete sb;
}
