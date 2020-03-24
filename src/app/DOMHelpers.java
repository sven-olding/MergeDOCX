package app;

import java.util.ArrayList;
import java.util.List;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.Text;

/**
 * This class contains static methods that help with collecting data from and
 * manipulating the MSWord XML-Structures.<br>
 * 
 */
public class DOMHelpers {
  /** contains the namespace uri for word components */
  public final static String NS_W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  /** contains the namespace uri for word relations */
  public final static String NS_R_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  /** contains the namespace uri for default xml components */
  public final static String NS_XML_URI = "http://www.w3.org/XML/1998/namespace";

  /** the tag name of bookmark start tags */
  public final static String NODE_BM_START = "bookmarkStart";

  /** the tag name of bookmark end tags */
  public final static String NODE_BM_END = "bookmarkEnd";

  /*
   * METHODS FOR CREATION
   */
  /**
   * Creates and returns an XML object of type 'w:r:' which contains the given
   * text in a 'w:t' subnode.<br>
   * Though a parentNode must be given, that is only used to retrieve the owner
   * document. The created node will NOT be added to the given parent node! <br>
   * <br>
   * 
   * @param text       content of the returned range object
   * @param parentNode the node where the result will be located in the end
   * @return a MS Word range XML object with the given text as content
   * @since 8.8.1
   */
  public final static Element createRangeWithText(String text, Node parentNode) {
    Document ownerDoc = parentNode.getOwnerDocument();

    Element newR = ownerDoc.createElementNS(NS_W_URI, "r");
    Element newT = ownerDoc.createElementNS(NS_W_URI, "t");
    Text newText = ownerDoc.createTextNode(text);

    newT.setAttributeNS(NS_XML_URI, "space", "preserve");
    newR.appendChild(newT);
    newT.appendChild(newText);

    return newR;
  }

  /**
   * Creates and returns an XML object of type 'w:r' which contains the given text
   * in a 'w:t' subnode. The text as list represents lines of text, which are
   * wrapped around by using 'w:br' tags.<br>
   * Though a parentNode must be given, that is only used to retrieve the owner
   * document. The created node will NOT be added to the given parent node! <br>
   * <br>
   * 
   * @param text       lines of content of the returned range object
   * @param parentNode the node where the result will be located in the end
   * @return a MS Word range XML object with the given text as content
   * @since 8.8.1
   */
  public final static Element createRangeWithText(List<String> text, Node parentNode) {
    Document ownerDoc = parentNode.getOwnerDocument();

    Element newR = ownerDoc.createElementNS(NS_W_URI, "r");

    if (text == null) {
      text = new ArrayList<String>();
    }

    if (text.isEmpty()) {
      text.add("");
    }

    for (int i = 0; i < text.size(); i++) {
      String line = text.get(i);
      Element newT = ownerDoc.createElementNS(NS_W_URI, "t");

      Text newText = ownerDoc.createTextNode(line);
      newT.setAttributeNS(NS_XML_URI, "space", "preserve");
      newT.appendChild(newText);

      newR.appendChild(newT);

      if (i < text.size() - 1) {
        Element br = ownerDoc.createElementNS(NS_W_URI, "br");
        newR.appendChild(br);
      }
    }

    return newR;
  }

  /**
   * Splits a paragraph at the corresponding bookmark start, where the
   * altChunk-tag for a mime file (richtext) must be inserted. The altChunk-tag
   * must not be positioned inside of a paragraph, it rather has to be a sibling
   * of it. So in this case, the paragraph is split in two, first one containing
   * bookmarkStart, other on containing bookmarkEnd and the altChunk is positioned
   * right between them.<br>
   * <br>
   * 
   * @param id            the id to be used as relation id of the new altChunk tag
   * @param bookmarkStart node to split the paragraph at
   * @since 8.8.1
   */
  public final static void splitParagraphForAltChunk(String id, Node bookmarkStart) {
    Node parentNode = bookmarkStart.getParentNode();
    Document ownerDoc = bookmarkStart.getOwnerDocument();

    Element newAltChunk = ownerDoc.createElementNS(NS_W_URI, "altChunk");
    newAltChunk.setAttributeNS(NS_R_URI, "id", id);

    if (isTagName(parentNode, "body")) {
      parentNode.appendChild(newAltChunk);
    } else if (isTagName(parentNode, "p")) {
      Node pPr = clonePreviousPPr(bookmarkStart);
      Element newP = ownerDoc.createElementNS(NS_W_URI, "p");
      NamedNodeMap parentAttibutes = parentNode.getAttributes();
      for (int i = 0; i < parentAttibutes.getLength(); i++) {
        Attr att = (Attr) parentAttibutes.item(i);
        newP.setAttributeNS(att.getNamespaceURI(), att.getName(), att.getValue());
      }

      if (pPr != null) {
        newP.appendChild(pPr);
      }

      Node next = bookmarkStart.getNextSibling();
      while (next != null) {
        newP.appendChild(next);
        next = bookmarkStart.getNextSibling();
      }

      insertNodeAfter(newAltChunk, parentNode);
      insertNodeAfter(newP, newAltChunk);
    } else {
      throw new IllegalStateException("Illegal position for altChunk. Parent: " + parentNode.getNodeName());
    }

  }

  /**
   * Creates and returns a 'w:bookmarkEnd' tag in the rare case that an end tag is
   * missing. It is constructed as end of the given start tag and thus uses that
   * tag's id.<br>
   * <br>
   * 
   * @param bookmarkStart tag to create the end tag for
   * @return the end tag for thr given start tag
   * @since 8.8.1
   */
  public final static Element createBookmarkEnd(Node bookmarkStart) {
    Document doc = bookmarkStart.getOwnerDocument();
    Element newEnd = doc.createElementNS(NS_W_URI, NODE_BM_END);
    newEnd.setAttributeNS(NS_W_URI, "id", getIdFromNode(bookmarkStart));
    return newEnd;
  }

  /*
   * METHODS FOR MOVING
   */
  /**
   * Moves the given nodeToInsert to the position after the reference node.<br>
   * This does <b>not</b> copy the given node, it is moved from its original
   * position. <br>
   * <br>
   * 
   * @param nodeToInsert the node to be moved to the target position
   * @param refNode      the nodeToInsert is positioned as next sibling to this
   *                     node
   * @since 8.8.1
   */
  public final static void insertNodeAfter(Node nodeToInsert, Node refNode) {
    Node nextNode = refNode.getNextSibling();
    if (nextNode == null) {
      refNode.getParentNode().appendChild(nodeToInsert);
    } else {
      refNode.getParentNode().insertBefore(nodeToInsert, nextNode);
    }
  }

  /*
   * METHODS FOR QUERYING
   */
  /**
   * Returns true, if the given node is a bookmarkEnd tag.<br>
   * <br>
   * 
   * @param node node to check
   * @return true, if the given node is a bookmarkEnd tag
   * @since 8.8.1
   */
  public final static boolean isBookmarkEnd(Node node) {
    return isTagName(node, NODE_BM_END);
  }

  /**
   * Returns true, if the given node is a bookmarkStart tag.<br>
   * <br>
   * 
   * @param node node to check
   * @return true, if the given node is a bookmarkStart tag
   * @since 8.8.1
   */
  public final static boolean isBookmarkStart(Node node) {
    return isTagName(node, NODE_BM_START);
  }

  /**
   * Returns true, if the given node is a tag of the given tag name.<br>
   * <br>
   * 
   * @param node node to check
   * @param tag  tag name to check
   * @return true, if the given node is a tag of the given tag name
   * @since 8.8.1
   */
  public final static boolean isTagName(Node node, String tag) {
    return node != null && node.getNodeName() != null && node.getNodeName().contains(tag);
  }

  /**
   * Retrieves the id attribute from the given node as text.<br>
   * <br>
   * 
   * @param node node to get the id from
   * @return the id attribute from the given node as text
   * @since 8.8.1
   */
  public final static String getIdFromNode(Node node) {
    return getAttributeFromNode(node, "id");
  }

  /**
   * Retrieves the name attribute from the given node as text.<br>
   * <br>
   * 
   * @param node node to get the name from
   * @return the name attribute from the given node as text
   * @since 8.8.1
   */
  public final static String getNameFromNode(Node node) {
    return getAttributeFromNode(node, "name");
  }

  /**
   * Retrieves the given attribute from the given node as text.<br>
   * <br>
   * 
   * @param node      node to get the attribute from
   * @param attribute name of the attribute to retrieve
   * @return the attribute from the given node as text
   * @since 8.8.1
   */
  public final static String getAttributeFromNode(Node node, String attribute) {
    String result = "";
    NamedNodeMap map = node.getAttributes();
    if (map != null) {
      Node idNode = map.getNamedItemNS(NS_W_URI, attribute);
      if (idNode != null) {
        result = idNode.getNodeValue();
      }
    }
    return result;
  }

  /**
   * Iterates through the XML DOM tree starting at the given node root and
   * collects and returns all nodes whose tags match the given name.<br>
   * <br>
   * 
   * @param root node to begin interation
   * @param name tag name of the nodes to be collected
   * @return a list of the collected nodes
   * @since 8.8.1
   */
  public final static List<Node> collectAllNodes(Node root, String name) {
    List<Node> result = new ArrayList<Node>();

    boolean repeat = true;

    while (repeat) {
      if (isTagName(root, name)) {
        result.add(root);
      }

      Node next = root.getFirstChild();
      boolean succ = next != null;
      if (succ) {
        root = next;
      } else {
        next = root.getNextSibling();
        succ = next != null;
        if (succ) {
          root = next;
        }
      }

      while (!succ && repeat) {
        next = root.getParentNode();
        succ = next != null;
        if (succ) {
          root = next;
        } else {
          repeat = false;
          break;
        }

        next = root.getNextSibling();
        succ = next != null;
        if (succ) {
          root = next;
        }
      }
    }

    return result;
  }

  /**
   * Searches backwards from the given node and returns a copy (including subtree)
   * of the first 'w:rPr' tag it encounters.<br>
   * This is needed for formatting fonts and sizes of a created MS Word range.<br>
   * <br>
   * 
   * @param node node to start search
   * @return a copy of the subtree containing the nearest 'w:rPr' tag
   * @since 8.8.1
   */
  public final static Node clonePreviousRPr(Node node) {
    return clonePreviousNodeOfType(node, "rPr");
  }

  /**
   * Searches backwards from the given node and returns a copy (including subtree)
   * of the first 'w:pPr' tag it encounters.<br>
   * This is needed for formatting a created MS Word paragraph.<br>
   * <br>
   * 
   * @param node node to start search
   * @return a copy of the subtree containing the nearest 'w:pPr' tag
   * @since 8.8.1
   */
  public final static Node clonePreviousPPr(Node node) {
    return clonePreviousNodeOfType(node, "pPr");
  }

  /**
   * Searches backwards from the given node and returns a copy (including subtree)
   * of the first tag with the given name it encounters.<br>
   * <br>
   * 
   * @param node     node to start search
   * @param nodeName name of the node to search for
   * @return a copy of the subtree containing the nearest tag with the given node
   *         name
   * @since 8.8.1
   */
  public final static Node clonePreviousNodeOfType(Node node, String nodeName) {
    Node result = null;
    boolean repeat = true;

    while (repeat) {
      if (DOMHelpers.isTagName(node, nodeName)) {
        result = node;
        break;
      }

      Node prevNode = node.getLastChild();
      boolean succ = prevNode != null;
      if (succ) {
        node = prevNode;
      } else {
        prevNode = node.getPreviousSibling();
        succ = prevNode != null;
        if (succ) {
          node = prevNode;
        }
      }

      while (!succ && repeat) {
        prevNode = node.getParentNode();
        succ = prevNode != null;
        if (succ) {
          node = prevNode;
        } else {
          repeat = false;
          break;
        }

        prevNode = node.getPreviousSibling();
        succ = prevNode != null;
        if (succ) {
          node = prevNode;
        }
      }
    }

    if (result != null) {
      result = result.cloneNode(true);
    }

    return result;
  }
}
