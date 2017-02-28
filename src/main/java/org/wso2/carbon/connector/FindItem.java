/*
*  Copyright (c) 2017, WSO2 Inc. (http://www.wso2.org) All Rights Reserved.
*
*  WSO2 Inc. licenses this file to you under the Apache License,
*  Version 2.0 (the "License"); you may not use this file except
*  in compliance with the License.
*  You may obtain a copy of the License at
*
*    http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing,
* software distributed under the License is distributed on an
* "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
* KIND, either express or implied.  See the License for the
* specific language governing permissions and limitations
* under the License.
*/
package org.wso2.carbon.connector;

import org.apache.axiom.om.OMAbstractFactory;
import org.apache.axiom.om.OMElement;
import org.apache.axiom.om.OMNamespace;
import org.apache.axiom.om.util.AXIOMUtil;
import org.apache.axiom.soap.SOAPBody;
import org.apache.axiom.soap.SOAPEnvelope;
import org.apache.axiom.soap.SOAPFactory;
import org.apache.axiom.soap.SOAPHeader;
import org.apache.axis2.AxisFault;
import org.apache.synapse.MessageContext;
import org.wso2.carbon.connector.core.AbstractConnector;
import org.wso2.carbon.connector.core.ConnectException;
import org.wso2.carbon.connector.core.util.ConnectorUtils;
import org.wso2.carbon.utils.xml.StringUtils;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLStreamException;
import javax.xml.transform.TransformerException;
import java.io.IOException;

import static org.wso2.carbon.connector.EWSUtils.populateItemShape;

/**
 * This Class used to Generate FindItem Operation SOAP Request
 */
public class FindItem extends AbstractConnector {
    OMNamespace type = EWSUtils.type;
    OMNamespace message = EWSUtils.message;
    SOAPFactory soapFactory = OMAbstractFactory.getSOAP11Factory();

    public void connect(MessageContext messageContext) throws ConnectException {
        SOAPEnvelope soapEnvelope = soapFactory.createSOAPEnvelope();
        soapEnvelope.declareNamespace(type);
        soapEnvelope.declareNamespace(message);
        try {
            soapEnvelope.addChild(populateSoapHeader(messageContext));
            soapEnvelope.addChild(populateBody(messageContext));
            messageContext.setEnvelope(soapEnvelope);
        } catch (XMLStreamException e) {
            String msg = "Couldn't convert Element Body";
            handleException(msg, e, messageContext);
        } catch (AxisFault axisFault) {
            String msg = "Couldn't set SOAPEnvelope to MessageContext";
            handleException(msg, axisFault, messageContext);
        } catch (TransformerException e) {
            String msg = "Couldn't transform message";
            handleException(msg, e, messageContext);
        } catch (IOException e) {
            String msg = "Couldn't locate xslt file";
            handleException(msg, e, messageContext);
        }
    }

    /**
     * Used to populate soap headers
     *
     * @param messageContext message context of request
     * @return Soap Header
     * @throws XMLStreamException
     * @throws TransformerException throws when
     */
    private SOAPHeader populateSoapHeader(MessageContext messageContext) throws XMLStreamException,
            TransformerException, IOException {
        SOAPHeader soapHeader = soapFactory.createSOAPHeader();
        EWSUtils.populateManagementRolesHeader(soapHeader, messageContext);
        EWSUtils.populateDateTimePrecisionHeader(soapHeader, messageContext);
        EWSUtils.populateTimeZoneContextHeader(soapHeader, messageContext);
        EWSUtils.populateRequestedServerVersionHeader(soapHeader, messageContext);
        EWSUtils.populateMailboxCulture(soapHeader, messageContext);
        EWSUtils.populateExchangeImpersonationHeader(soapHeader, messageContext);
        return soapHeader;
    }

    /**
     * Used to populate soap body
     *
     * @param messageContext message context of request
     * @return Soap Body
     * @throws XMLStreamException
     * @throws TransformerException throws when
     */
    private SOAPBody populateBody(MessageContext messageContext) throws XMLStreamException, TransformerException,
            IOException {
        SOAPBody soapBody = soapFactory.createSOAPBody();
        OMElement findItemElement = soapFactory.createOMElement(EWSConstants.FIND_ITEM_ELEMENT, message);
        EWSUtils.setValueToXMLAttribute(messageContext, findItemElement, EWSConstants.TRAVERSAL, EWSConstants
                .TRAVERSAL_ELEMENT);
        findItemElement.addChild(populateItemShape(messageContext));
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.INDEXED_PAGE_ITEM_VIEW, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.FRACTIONAL_PAGE_ITEM_VIEW,
                message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants
                .SEEK_TO_CONDITION_PAGE_ITEM_VIEW, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.CALENDAR_VIEW, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.CONTACTS_VIEW, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.GROUP_BY, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.DISTINGUISHED_GROUP_BY, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.RESTRICTION, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.SORT_ORDER, message);
        EWSUtils.populateDirectElements(messageContext, findItemElement, EWSConstants.PARENT_FOLDER_IDS, message);
        OMElement queryStringElement = soapFactory.createOMElement("QueryString", message);
        String queryString = (String) ConnectorUtils.lookupTemplateParamater(messageContext, EWSConstants.QUERY_STRING);
        if (!StringUtils.isEmpty(queryString)) {
            OMElement omElement = AXIOMUtil.stringToOM(queryString);
            OMElement resetCache = omElement.getFirstChildWithName(new QName("ResetCache"));
            OMElement returnHighlightTerms = omElement.getFirstChildWithName(new QName("ReturnHighlightTerms"));
            OMElement returnDeletedItems = omElement.getFirstChildWithName(new QName("ReturnDeletedItems"));
            OMElement body = omElement.getFirstChildWithName(new QName(EWSConstants.BODY_ELEMENT));
            if (resetCache != null) {
                queryStringElement.addAttribute("ResetCache", resetCache.getText(), null);
            }
            if (returnHighlightTerms != null) {
                queryStringElement.addAttribute("ReturnHighlightTerms", returnHighlightTerms.getText(), null);
            }
            if (returnDeletedItems != null) {
                queryStringElement.addAttribute("ReturnDeletedItems", returnDeletedItems.getText(), null);
            }
            if (body != null) {
                queryStringElement.setText(body.getText());
            }
            findItemElement.addChild(queryStringElement);
        }
        soapBody.addChild(findItemElement);
        return soapBody;
    }
}
