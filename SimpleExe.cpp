//Armando Alvarez
//CSC546
//Spring 2025
//Autodesk Inventor to Hipair Interface

#include "stdafx.h"
#include <fstream>
#include <string>
#include <iomanip>

static HRESULT GetInventorInformation(std::ofstream& out);
static void parseGeom(CComPtr<Document> pDoc, std::ofstream& out);

int _tmain(int argc, _TCHAR* argv[])
{
    HRESULT Result = NOERROR;

    Result = CoInitialize(NULL);

    std::ofstream out("C:/Users/littl/source/repos/hipair/assembly.txt");
    if (!out.is_open()) {
        _tprintf_s(_T("Could not open file.\n"));
        return 1;
    }

    if (SUCCEEDED(Result))
        Result = GetInventorInformation(out);

    CoUninitialize();
    out.close();
    return 0;
}

static HRESULT GetInventorInformation(std::ofstream& out)
{
    HRESULT Result = NOERROR;

    CLSID InvAppClsid;
    Result = CLSIDFromProgID(L"Inventor.Application", &InvAppClsid);
    if (FAILED(Result)) return Result;

    CComPtr<IUnknown> pInvAppUnk;
    Result = ::GetActiveObject(InvAppClsid, NULL, &pInvAppUnk);
    if (FAILED(Result))
    {
        _tprintf_s(_T("*** Could not get hold of an active Inventor application ***\n"));
        return Result;
    }

    CComPtr<Application> pInvApp;
    Result = pInvAppUnk->QueryInterface(__uuidof(Application), (void**)&pInvApp);
    if (FAILED(Result)) return Result;

    CComPtr<Document> pDoc;
    Result = pInvApp->get_ActiveDocument(&pDoc);
    if (SUCCEEDED(Result) && pDoc)
    {
        parseGeom(pDoc, out);
    }
    else
    {
        _tprintf_s(_T("*** No active document found. ***\n"));
    }

    return Result;
}

static void parseGeom(CComPtr<Document> pDoc, std::ofstream& out)
{
    if (!pDoc)
        return;

    DocumentTypeEnum docType;
    HRESULT hr = pDoc->get_DocumentType(&docType);
    if (FAILED(hr))
        return;

    if (docType == kAssemblyDocumentObject)
    {
        _tprintf_s(_T("Assembly is active\n"));

        out << "1 1 2 1 0 0\n\n";

        CComQIPtr<AssemblyDocument> pAssemblyDoc(pDoc);
        if (!pAssemblyDoc)
            return;

        CComPtr<AssemblyComponentDefinition> pAssemblyDef;
        hr = pAssemblyDoc->get_ComponentDefinition(&pAssemblyDef);
        if (FAILED(hr))
            return;

        CComPtr<ComponentOccurrences> pOccurrences;
        hr = pAssemblyDef->get_Occurrences(&pOccurrences);
        if (FAILED(hr))
            return;

        long occCount = 0;
        pOccurrences->get_Count(&occCount);
        _tprintf_s(_T("Count of Assembly occurrences: %d\n"), occCount);

        for (long i = 1; i <= occCount; ++i)
        {
            CComPtr<ComponentOccurrence> pOccur;
            hr = pOccurrences->get_Item(i, &pOccur);
            if (FAILED(hr) || !pOccur)
                continue;

            CComBSTR bstrName;
            hr = pOccur->get_Name(&bstrName);
            if (SUCCEEDED(hr))
                _tprintf_s(_T("%d component name: %ls\n"), i, bstrName);

            CComPtr<ComponentDefinition> pCompDef;
            hr = pOccur->get_Definition(&pCompDef);
            if (SUCCEEDED(hr) && pCompDef)
            {
                CComPtr<IDispatch> pSubDocDisp;
                hr = pCompDef->get_Document(&pSubDocDisp);
                if (SUCCEEDED(hr) && pSubDocDisp)
                {
                    CComQIPtr<Document> pSubDoc(pSubDocDisp);
                    if (pSubDoc)
                    {
                        out << "Part 1 255 0 0 0  3.0 1.0 1.0 0.0 -3.14159 3.14159 0.0\n";
                        out << "1 0.0 1.0\n";
                        out << "1 4\n";

                        parseGeom(pSubDoc, out);
                        out << "1 0.0\n\n";
 
                    }
                }
            }
        }
     out << "0 1 0.001\n\n";
     out << "2\n";
     out << "7.0 1 0 -1.0\n";
     out << "5.0 1 1 3.0\n";
     out << "0 3\n\n";
     out << "0 0 -5 5 -5 5";
    }
    else if (docType == kPartDocumentObject)
    {
        _tprintf_s(_T("This is a part document.\n"));

        CComQIPtr<PartDocument> pPartDoc(pDoc);
        if (!pPartDoc)
            return;

        CComPtr<PartComponentDefinition> pPartDef;
        hr = pPartDoc->get_ComponentDefinition(&pPartDef);
        if (FAILED(hr))
            return;

        CComPtr<PlanarSketches> pSketches;
        hr = pPartDef->get_Sketches(&pSketches);
        if (FAILED(hr) || !pSketches)
            return;

        long sketchCount = 0;
        pSketches->get_Count(&sketchCount);
        _tprintf_s(_T("Number of Sketches: %d\n"), sketchCount);

        if (sketchCount > 0)
        {
            CComPtr<PlanarSketch> pSketch;
            hr = pSketches->get_Item(CComVariant(1), &pSketch);
            if (FAILED(hr) || !pSketch)
                return;

            CComPtr<SketchEntitiesEnumerator> pEntities;
            hr = pSketch->get_SketchEntities(&pEntities);
            if (FAILED(hr) || !pEntities)
                return;

            long entityCount = 0;
            pEntities->get_Count(&entityCount);
            _tprintf_s(_T("Number of sketch entities: %d\n"), entityCount);

            for (long j = entityCount; j >= 1; --j)
            {
                CComPtr<SketchEntity> pEntity;
                hr = pEntities->get_Item(j, &pEntity);
                if (FAILED(hr) || !pEntity)
                    continue;

                ObjectTypeEnum type;
                pEntity->get_Type(&type);

                if (type == kSketchLineObject)
                {
                    CComQIPtr<SketchLine> pLine(pEntity);
                    if (!pLine)
                        continue;

                    double length;
                    pLine->get_Length(&length);
                    _tprintf_s(_T("Line %d length: %.6f cm\n"), j, length);

                    CComPtr<SketchPoint> pStartPt, pEndPt;
                    pLine->get_StartSketchPoint(&pStartPt);
                    pLine->get_EndSketchPoint(&pEndPt);

                    if (pStartPt && pEndPt)
                    {
                        CComPtr<Point2d> pStartGeom, pEndGeom;
                        pStartPt->get_Geometry(&pStartGeom);
                        pEndPt->get_Geometry(&pEndGeom);

                        if (pStartGeom && pEndGeom)
                        {
                            double x1, y1, x2, y2;
                            pStartGeom->get_X(&x1);
                            pStartGeom->get_Y(&y1);
                            pEndGeom->get_X(&x2);
                            pEndGeom->get_Y(&y2);

                            _tprintf_s(_T("\tLine Start Point: (%.6f, %.6f), End Point: (%.6f, %.6f)\n"), x1, y1, x2, y2);
                            out << "0 0 "
                                << std::fixed << std::setprecision(6)
                                << x2 << " " << y2 << " " << x1 << " " << y1 << " 1\n";
                        }
                    }
                }
            }
        }
    }
}
