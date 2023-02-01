#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Hyperlinks.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveHyperlinks.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get all hyperlinks
	vector<Field*> hyperlinks = FindAllHyperlinks(doc);

	//Flatten all hyperlinks
	for (int i = hyperlinks.size() - 1; i >= 0; i--)
	{
		FlattenHyperlinks(hyperlinks[i]);
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

vector<Field*> FindAllHyperlinks(Document* document)
{
	vector<Field*> hyperlinks;
	//Iterate through the items in the sections to find all hyperlinks
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		Section* section = document->GetSections()->GetItem(i);
		int secBodyChildCount = section->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < secBodyChildCount; j++)
		{
			DocumentObject* childObj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (childObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				int paraChildCount = (dynamic_cast<Paragraph*>(childObj))->GetChildObjects()->GetCount();
				for (int k = 0; k < paraChildCount; k++)
				{
					DocumentObject* paraObj = (dynamic_cast<Paragraph*>(childObj))->GetChildObjects()->GetItem(k);
					if (paraObj->GetDocumentObjectType() == DocumentObjectType::Field)
					{
						Field* field = dynamic_cast<Field*>(paraObj);
						if (field->GetType() == FieldType::FieldHyperlink)
						{
							hyperlinks.push_back(field);
						}
					}
				}
			}
		}
	}
	return hyperlinks;
}

void FlattenHyperlinks(Field* field)
{
	int ownerParaIndex = field->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetOwnerParagraph());
	int fieldIndex = field->GetOwnerParagraph()->GetChildObjects()->IndexOf(field);
	Paragraph* sepOwnerPara = field->GetSeparator()->GetOwnerParagraph();
	int sepOwnerParaIndex = field->GetSeparator()->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetSeparator()->GetOwnerParagraph());
	int sepIndex = field->GetSeparator()->GetOwnerParagraph()->GetChildObjects()->IndexOf(field->GetSeparator());
	int endIndex = field->GetEnd()->GetOwnerParagraph()->GetChildObjects()->IndexOf(field->GetEnd());
	int endOwnerParaIndex = field->GetEnd()->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetEnd()->GetOwnerParagraph());

	FormatFieldResultText(field->GetSeparator()->GetOwnerParagraph()->GetOwnerTextBody(), sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

	field->GetEnd()->GetOwnerParagraph()->GetChildObjects()->RemoveAt(endIndex);

	for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--)
	{
		if (i == sepOwnerParaIndex && i == ownerParaIndex)
		{
			for (int j = sepIndex; j >= fieldIndex; j--)
			{
				field->GetOwnerParagraph()->GetChildObjects()->RemoveAt(j);

			}
		}
		else if (i == ownerParaIndex)
		{
			for (int j = field->GetOwnerParagraph()->GetChildObjects()->GetCount() - 1; j >= fieldIndex; j--)
			{
				field->GetOwnerParagraph()->GetChildObjects()->RemoveAt(j);
			}

		}
		else if (i == sepOwnerParaIndex)
		{
			for (int j = sepIndex; j >= 0; j--)
			{
				sepOwnerPara->GetChildObjects()->RemoveAt(j);
			}
		}
		else
		{
			field->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->RemoveAt(i);
		}
	}
}

void FormatFieldResultText(Body* ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
{
	for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
	{
		Paragraph* para = dynamic_cast<Paragraph*>(ownerBody->GetChildObjects()->GetItem(i));
		if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
		{
			for (int j = sepIndex + 1; j < endIndex; j++)
			{
				FormatText(dynamic_cast<TextRange*>(para->GetChildObjects()->GetItem(j)));
			}

		}
		else if (i == sepOwnerParaIndex)
		{
			for (int j = sepIndex + 1; j < para->GetChildObjects()->GetCount(); j++)
			{
				FormatText(dynamic_cast<TextRange*>(para->GetChildObjects()->GetItem(j)));
			}
		}
		else if (i == endOwnerParaIndex)
		{
			for (int j = 0; j < endIndex; j++)
			{
				FormatText(dynamic_cast<TextRange*>(para->GetChildObjects()->GetItem(j)));
			}
		}
		else
		{
			for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
			{
				FormatText(dynamic_cast<TextRange*>(para->GetChildObjects()->GetItem(j)));
			}
		}
	}
}

void FormatText(TextRange* tr)
{
	//Set the text color to black
	tr->GetCharacterFormat()->SetTextColor(Color::GetBlack());
	//Set the text underline style to none
	tr->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::None);
}
