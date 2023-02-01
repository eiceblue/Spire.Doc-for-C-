#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Hyperlinks.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FindHyperlinks.txt";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Create a hyperlink list
	vector<Field*> hyperlinks;
	wstring hyperlinksText = L"";
	//Iterate through the items in the sections to find all hyperlinks
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* section = doc->GetSections()->GetItem(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			DocumentObject* docObj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (docObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				Paragraph* para = dynamic_cast<Paragraph*>(docObj);
				for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
				{
					DocumentObject* obj = para->GetChildObjects()->GetItem(k);
					if (obj->GetDocumentObjectType() == DocumentObjectType::Field)
					{
						Field* field = dynamic_cast<Field*>(obj);
						if (field->GetType() == FieldType::FieldHyperlink)
						{
							hyperlinks.push_back(field);
							//Get the hyperlink text
							wstring text = field->GetFieldText();
							hyperlinksText.append(text.append(L"\r\n"));
						}
					}
				}
			}
		}
	}

	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		if (strcmp(typeid(doc->GetSections()->GetItem(i)).name(), typeid(Section).name()))
			Section sec = *doc->GetSections()->GetItem(i);
	}

	//Save the text of all hyperlinks to TXT File and launch it
	wofstream write(outputFile);
	write << hyperlinksText;
	write.close();
	//File::WriteAllText(outputFile.c_str(), hyperlinksText);
	doc->Close();
	delete doc;
}
