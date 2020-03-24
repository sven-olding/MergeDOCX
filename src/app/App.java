package app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URI;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.PackageHelper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.w3c.dom.Node;

public class App {
  public static void main(final String[] args) {
    File mainFile = new File("./docx/Main.docx");
    File subFile = new File("./docx/Sub.docx");

    try (InputStream inputStreamMainFile = new FileInputStream(mainFile);
        InputStream inputStreamSubFile = new FileInputStream(subFile);
        FileOutputStream fileOutputStream = new FileOutputStream(new File("./docx/Result.docx"));) {
      XWPFDocument wordDocMain = new XWPFDocument(inputStreamMainFile);
      XWPFDocument wordDocSub = new XWPFDocument(inputStreamSubFile);

      XWPFDocument wordDocResult = mergeDocuments(wordDocMain, wordDocSub);

      wordDocResult.write(fileOutputStream);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static XWPFDocument mergeDocuments(XWPFDocument sourceDoc, XWPFDocument docToAdd) throws Exception {
    try {
      stripUnneededBookmarks(docToAdd);
    } catch (Throwable t) {
      // ignore
    }

    OPCPackage mergePkg1 = PackageHelper.clone(sourceDoc.getPackage(), PackageHelper.createTempFile());

    XWPFDocument mergeDoc1 = new XWPFDocument(mergePkg1);
    mergeDoc1.getDocument().unsetBody();

    CTBody mainBody = sourceDoc.getDocument().getBody();

    XmlOptions optionsOuter = new XmlOptions();
    optionsOuter.setSaveInner();
    CTBody mergeBody1 = mergeDoc1.getDocument().addNewBody();
    mergeBody1.set(mainBody);

    CTBody addBody = docToAdd.getDocument().getBody();

    String strAddBody1 = addBody.xmlText(optionsOuter);

    addNewBodyAsBody(mergeBody1, strAddBody1);

    return mergeDoc1;
  }

  /**
   * Removes the '_GoBack' bookmarks from the given XWPFDocument object. These
   * bookmarks are completly invisible to the user and only used by Word for
   * change tracking, so they can (must) be removed without any difficulties.<br>
   * <br>
   * 
   * @param wordDoc the XWPFDocument to remove the bookmarks from
   */
  private static void stripUnneededBookmarks(XWPFDocument wordDoc) {
    List<Node> startNodes = DOMHelpers.collectAllNodes(wordDoc.getDocument().getDomNode(), DOMHelpers.NODE_BM_START);
    List<Node> endNodes = DOMHelpers.collectAllNodes(wordDoc.getDocument().getDomNode(), DOMHelpers.NODE_BM_END);

    for (Node start : startNodes) {
      String bmName = DOMHelpers.getNameFromNode(start);
      if (bmName.equalsIgnoreCase("_GoBack")) {
        String startId = DOMHelpers.getIdFromNode(start);
        for (Node end : endNodes) {
          if (DOMHelpers.getIdFromNode(end).equals(startId)) {
            end.getParentNode().removeChild(end);
          }
        }
        start.getParentNode().removeChild(start);
      }
    }
  }

  /**
   * Internal helper method for merging documents.<br>
   * <br>
   * 
   * @param mainBody    CTBody object where the other object is appended to
   * @param strAddBody1 CTBody object which is appended to the other object
   * @throws Exception if anything goes wrong
   */
  private static void addNewBodyAsBody(CTBody mainBody, String strAddBody1) throws Exception {
    String strMainBody = mainBody.xmlText();

    HashMap<String, String> targetPrefixParts = new HashMap<String, String>();

    String mainPrefix = strMainBody.substring(0, strMainBody.indexOf(">") + 1);
    String[] mainPrefixPartsArray = mainPrefix.split(" ");
    ArrayList<String> mainPrefixParts = new ArrayList<String>(Arrays.asList(mainPrefixPartsArray));
    // remove <xml-fragment (first element) & remove ">" from last tag
    mainPrefixParts.remove(0);
    String lastElement = mainPrefixParts.remove(mainPrefixParts.size() - 1);
    if (lastElement.endsWith(">")) {
      lastElement = lastElement.substring(0, lastElement.length() - 1);
    }
    mainPrefixParts.add(lastElement);

    for (String pt : mainPrefixParts) {
      String[] splt = pt.split("=");
      targetPrefixParts.put(splt[0], splt[1]);
    }

    String mainPart = strMainBody.substring(strMainBody.indexOf(">") + 1, strMainBody.lastIndexOf("<"));
    String sufix = strMainBody.substring(strMainBody.lastIndexOf("<"));

    String addPrefix = strAddBody1.substring(0, strAddBody1.indexOf(">") + 1);
    String[] addPrefixPartsArray = addPrefix.split(" ");
    ArrayList<String> addPrefixParts = new ArrayList<String>(Arrays.asList(addPrefixPartsArray));
    // remove <xml-fragment (first element) & remove ">" from last tag
    addPrefixParts.remove(0);
    if (addPrefixParts.size() > 0) {
      lastElement = addPrefixParts.remove(addPrefixParts.size() - 1);
      if (lastElement.endsWith(">")) {
        lastElement = lastElement.substring(0, lastElement.length() - 1);
      }
    }
    addPrefixParts.add(lastElement);
    for (String pt : addPrefixParts) {
      String[] splt = pt.split("=");
      targetPrefixParts.put(splt[0], splt[1]);
    }

    String prefix = "<xml-fragment";
    for (String key : targetPrefixParts.keySet()) {
      prefix += " " + key + "=" + targetPrefixParts.get(key);
    }
    prefix += ">";

    String addPart1 = strAddBody1;
    if (addPart1.startsWith("<xml-fragment")) {
      addPart1 = addPart1.substring(strAddBody1.indexOf(">") + 1, strAddBody1.lastIndexOf("<"));
    }

    // correct bookmark ids
    // first scan ids on main part
    int nextId = 0;
    int ind1 = mainPart.indexOf("<w:bookmarkStart");
    while (ind1 > -1) {
      int ind2 = mainPart.indexOf("/>", ind1);
      if (ind2 > -1) {
        ind1 = mainPart.indexOf("id=\"", ind1);
        if (ind1 > -1) {
          ind1 += 4;
          ind2 = mainPart.indexOf("\"", ind1);
          String id = mainPart.substring(ind1, ind2);
          try {
            nextId = Integer.parseInt(id) + 1;
          } catch (NumberFormatException e) {
            // MUST NOT HAPPEN
          }
        }
      }
      ind1 = mainPart.indexOf("<w:bookmarkStart", ind1 + 1);
    }

    // then correct ids in addPart1
    ind1 = addPart1.indexOf("<w:bookmarkStart");
    while (ind1 > -1) {
      String currentId = "";
      int ind2 = addPart1.indexOf("/>", ind1);
      if (ind2 > -1) {
        ind1 = addPart1.indexOf("id=\"", ind1);
        if (ind1 > -1) {
          ind1 += 4;
          ind2 = addPart1.indexOf("\"", ind1);
          currentId = addPart1.substring(ind1, ind2);

        }
      }

      // find corresponding bookmarkEnd
      int ind3 = addPart1.indexOf("<w:bookmarkEnd", ind1);
      while (ind3 > -1) {
        String currentEndId = "";
        int ind4 = addPart1.indexOf("/>", ind3);
        if (ind4 > -1) {
          ind3 = addPart1.indexOf("id=\"", ind3);
          if (ind3 > -1) {
            ind3 += 4;
            ind4 = addPart1.indexOf("\"", ind3);
            currentEndId = addPart1.substring(ind3, ind4);
            if (currentId.equals(currentEndId)) {
              // change Ids of start and end to next id
              String p1 = addPart1.substring(0, ind1);
              String p2 = addPart1.substring(ind2, ind3);
              String p3 = addPart1.substring(ind4);
              addPart1 = p1 + nextId + p2 + nextId + p3;
              nextId++;
              break;
            }
          }
        }
      }

      ind1 = addPart1.indexOf("<w:bookmarkStart", ind1 + 1);
    }

    String fullXml = prefix + mainPart + addPart1 + sufix;

    CTBody makeBody = CTBody.Factory.parse(fullXml);

    // the new body must only contain one SectPtr part; we'll keep the first one we
    // find
    XmlCursor cur = makeBody.newCursor();
    boolean foundOneAlready = false;
    if (cur.toFirstChild()) {
      while (cur.toNextSibling()) {
        if (cur.getObject() instanceof CTSectPr) {
          if (foundOneAlready) {
            cur.removeXml();
          } else {
            foundOneAlready = true;
          }
        }
      }
    }

    mainBody.set(makeBody);
  }
}