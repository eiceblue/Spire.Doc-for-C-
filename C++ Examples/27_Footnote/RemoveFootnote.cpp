#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Footnote.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveFootnote.docx";

	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());
	Section* section = document->GetSections()->GetItem(0);

	//traverse paragraphs in the section and find the footnote
	for (int k = 0; k < section->GetParagraphs()->GetCount(); k++)
	{
		Paragraph* para = section->GetParagraphs()->GetItem(k);
		int index = -1;
		for (int i = 0, cnt = para->GetChildObjects()->GetCount(); i < cnt; i++)
		{
			ParagraphBase* pBase = dynamic_cast<ParagraphBase*>(para->GetChildObjects()->GetItem(i));
			if (dynamic_cast<Footnote*>(pBase) != nullptr)
			{
				index = i;
				break;
			}
		}

		if (index > -1)
		{
			//remove the footnote
			para->GetChildObjects()->RemoveAt(index);
		}
	}

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
