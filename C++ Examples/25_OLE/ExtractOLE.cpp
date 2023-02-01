#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"OLEs.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile_pdf = output_path + L"ExtractOLE.pdf";
	wstring outputFile_xls = output_path + L"ExtractOLE.xls";
	wstring outputFile_pptx = output_path + L"ExtractOLE.pptx";

	//Create document and load file from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Traverse through all sections of the word document    
	for (int s = 0; s < doc->GetSections()->GetCount(); s++)
	{
		Section* sec = doc->GetSections()->GetItem(s);
		//Traverse through all Child Objects in the GetBody() of each section
		for (int i = 0; i < sec->GetBody()->GetChildObjects()->GetCount(); i++)
		{
			DocumentObject* obj = sec->GetBody()->GetChildObjects()->GetItem(i);
			//find the paragraph
			if (dynamic_cast<Paragraph*>(obj) != nullptr)
			{
				Paragraph* par = dynamic_cast<Paragraph*>(obj);
				for (int j = 0; j < par->GetChildObjects()->GetCount(); j++)
				{
					DocumentObject* o = par->GetChildObjects()->GetItem(j);
					//check whether the object is OLE
					if (o->GetDocumentObjectType() == DocumentObjectType::OleObject)
					{
						DocOleObject* Ole = dynamic_cast<DocOleObject*>(o);
						wstring s = Ole->GetObjectType();

						//check whether the object type is "Acrobat.Document.11"
						if (s == L"AcroExch.Document.DC")
						{
							//write the data of OLE into file										
							ofstream pdf_file(outputFile_pdf, ios::out | ofstream::binary);
							vector<byte> native_data = Ole->GetNativeData();
							pdf_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
							pdf_file.close();
							//File::WriteAllBytes(outputFile_pdf.c_str(), Ole->GetNativeData());
						}

						//check whether the object type is "Excel.Sheet.8"
						else if (s == L"Excel.Sheet.8")
						{
							ofstream xls_file(outputFile_xls, ios::out | ofstream::binary);
							vector<byte> native_data = Ole->GetNativeData();
							xls_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
							xls_file.close();
							//File::WriteAllBytes(outputFile_xls.c_str(), Ole->GetNativeData());										
						}

						//check whether the object type is "PowerPoint.Show.12"
						else if (s == L"PowerPoint.Show.12")
						{
							ofstream pptx_file(outputFile_pptx, ios::out | ofstream::binary);
							vector<byte> native_data = Ole->GetNativeData();
							pptx_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
							pptx_file.close();
							//File::WriteAllBytes(outputFile_pptx.c_str(), Ole->GetNativeData());
						}
					}
				}
			}
		}
	}
	doc->Close();
	delete doc;
}