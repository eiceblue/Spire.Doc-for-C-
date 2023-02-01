#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SetCaptionWithChapterNumber.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetCaptionWithChapterNumber.docx";

	//Create word document
	Document* document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	//Get the first section
	Section* section = document->GetSections()->GetItem(0);
	//Label name
	wstring name = L"Caption ";
	for (int i = 0; i < section->GetBody()->GetParagraphs()->GetCount(); i++)
	{
		for (int j = 0; j < section->GetBody()->GetParagraphs()->GetItem(i)->GetChildObjects()->GetCount(); j++)
		{
			if (dynamic_cast<DocPicture*>(section->GetBody()->GetParagraphs()->GetItem(i)->GetChildObjects()->GetItem(j)) != nullptr)
			{
				DocPicture* pic1 = dynamic_cast<DocPicture*>(section->GetBody()->GetParagraphs()->GetItem(i)->GetChildObjects()->GetItem(j));
				Body* body = dynamic_cast<Body*>(pic1->GetOwnerParagraph()->GetOwner());
				if (body != nullptr)
				{
					int imageIndex = body->GetChildObjects()->IndexOf(pic1->GetOwnerParagraph());
					//Create a new paragraph
					Paragraph* para = new Paragraph(document);
					//Set label
					para->AppendText(name.c_str());

					//Add caption
					Field* field1 = para->AppendField(L"test", FieldType::FieldStyleRef);
					//Chapter number
					field1->SetCode(L" STYLEREF 1 \\s ");
					//Chapter delimiter
					para->AppendText(L" - ");

					//Add picture sequence number
					SequenceField* field2 = dynamic_cast<SequenceField*>(para->AppendField(name.c_str(), FieldType::FieldSequence));
					field2->SetCaptionName(name.c_str());
					field2->SetNumberFormat(CaptionNumberingFormat::Number);
					body->GetParagraphs()->Insert(imageIndex + 1, para);

				}
			}
		}
	}
	//Set update fields
	document->SetIsUpdateFields(true);
	//Save the result file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
