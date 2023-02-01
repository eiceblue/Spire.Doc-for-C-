#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveContentControls.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveContentControls.docx";

	//Load document from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Loop through sections
	for (int s = 0; s < doc->GetSections()->GetCount(); s++)
	{
		Section* section = doc->GetSections()->GetItem(s);
		for (int i = 0; i < section->GetBody()->GetChildObjects()->GetCount(); i++)
		{
			//Loop through contents in paragraph
			if (dynamic_cast<Paragraph*>(section->GetBody()->GetChildObjects()->GetItem(i)) != nullptr)
			{
				Paragraph* para = dynamic_cast<Paragraph*>(section->GetBody()->GetChildObjects()->GetItem(i));
				for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
				{
					//Find the StructureDocumentTagInline
					if (dynamic_cast<StructureDocumentTagInline*>(para->GetChildObjects()->GetItem(j)) != nullptr)
					{
						StructureDocumentTagInline* sdt = dynamic_cast<StructureDocumentTagInline*>(para->GetChildObjects()->GetItem(j));
						//Remove the content control from paragraph
						para->GetChildObjects()->Remove(sdt);
						j--;
					}
				}
			}
			if (dynamic_cast<StructureDocumentTag*>(section->GetBody()->GetChildObjects()->GetItem(i)) != nullptr)
			{
				StructureDocumentTag* sdt = dynamic_cast<StructureDocumentTag*>(section->GetBody()->GetChildObjects()->GetItem(i));
				section->GetBody()->GetChildObjects()->Remove(sdt);
				i--;
			}
		}
	}

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}
