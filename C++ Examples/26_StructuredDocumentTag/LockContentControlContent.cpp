#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LockContentControlContent.docx";

	wstring htmlString;
	htmlString.append(L"<table style=\"width: 100 % \">")
		.append(L"<tr><th> Number </th><th> Name </th ><th>Age</th ></tr>")
		.append(L"<tr><td> 1 </td><td> Smith </td><td> 50 </td></tr>")
		.append(L"<tr> <td> 2 </td><td> Jackson </td><td> 94 </td> </tr>")
		.append(L"</table>");

	Document* doc = new Document();
	Section* section = doc->AddSection();
	Paragraph* paragraph = section->AddParagraph();
	paragraph->AppendHTML(htmlString.c_str());

	//Create StructureDocumentTag for document
	StructureDocumentTag* sdt = new StructureDocumentTag(doc);
	Section* section2 = doc->AddSection();
	section2->GetBody()->GetChildObjects()->Add(sdt);

	//Specify the type
	sdt->GetSDTProperties()->SetSDTType(SdtType::RichText);

	for (int i = 0; i < section->GetBody()->GetChildObjects()->GetCount(); i++)
	{
		DocumentObject* obj = section->GetBody()->GetChildObjects()->GetItem(i);
		if (obj->GetDocumentObjectType() == DocumentObjectType::Table)
		{
			sdt->GetSDTContent()->GetChildObjects()->Add(obj->Clone());
		}
	}

	//Lock content
	sdt->GetSDTProperties()->SetLockSettings(LockSettingsType::ContentLocked);

	doc->GetSections()->Remove(section);

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}
