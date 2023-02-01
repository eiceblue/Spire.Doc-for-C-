#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Template_Docx_4.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemovePageBreaks.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Traverse every paragraph of the first section of the document.
	for (int j = 0; j < document->GetSections()->GetItem(0)->GetParagraphs()->GetCount(); j++)
	{
		Paragraph* p = document->GetSections()->GetItem(0)->GetParagraphs()->GetItem(j);

		//Traverse every child object of a paragraph.
		for (int i = 0; i < p->GetChildObjects()->GetCount(); i++)
		{
			DocumentObject* obj = p->GetChildObjects()->GetItem(i);

			//Find the page break object.
			if (obj->GetDocumentObjectType() == DocumentObjectType::Break)
			{
				Break* b = dynamic_cast<Break*>(obj);

				//Remove the page break object from paragraph.
				p->GetChildObjects()->Remove(b);
			}
		}
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
