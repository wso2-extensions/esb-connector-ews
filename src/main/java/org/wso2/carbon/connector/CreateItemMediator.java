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

import static org.wso2.carbon.connector.EWSUtils.populateDirectElements;
import static org.wso2.carbon.connector.EWSUtils.populateSaveItemFolderIdElement;

/**
 * used to generate CreateItem Soap Request
 */
public class CreateItemMediator extends AbstractConnector {
    private OMNamespace type = EWSUtils.type;
    private OMNamespace message = EWSUtils.message;
    private SOAPFactory soapFactory = OMAbstractFactory.getSOAP11Factory();

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
            log.error(msg, e);
            throw new ConnectException(e, msg);
        } catch (AxisFault axisFault) {
            String msg = "Couldn't set SOAPEnvelope to MessageContext";
            log.error(msg, axisFault);
            throw new ConnectException(axisFault, msg);
        } catch (TransformerException e) {
            String msg = "Couldn't transform message";
            log.error(msg, e);
            throw new ConnectException(e, msg);
        }

    }

    /**
     * Used to populate soap headers
     * @param messageContext message context of request
     * @return Soap Header
     * @throws XMLStreamException
     * @throws TransformerException throws when 
     */
    private SOAPHeader populateSoapHeader(MessageContext messageContext) throws XMLStreamException,
            TransformerException {
        SOAPHeader soapHeader = soapFactory.createSOAPHeader();
        EWSUtils.populateTimeZoneContextHeader(soapHeader, messageContext);
        EWSUtils.populateRequestedServerVersionHeader(soapHeader, messageContext);
        EWSUtils.populateMailboxCulture(soapHeader, messageContext);
        EWSUtils.populateExchangeImpersonationHeader(soapHeader, messageContext);
        return soapHeader;
    }

    /**
     * Used to populate soap body
     * @param messageContext message context of request
     * @return Soap Body
     * @throws XMLStreamException
     * @throws TransformerException
     */
    private SOAPBody populateBody(MessageContext messageContext) throws XMLStreamException, TransformerException {
        SOAPBody soapBody = soapFactory.createSOAPBody();
        OMElement createItemElement = soapFactory.createOMElement(EWSConstants.CREATE_ITEM_ELEMENT, message);
        EWSUtils.setValueToXMLAttribute(messageContext, createItemElement, EWSConstants.MESSAGE_DISPOSITION,
                EWSConstants.MESSAGE_DISPOSITION_ELEMENT);
        EWSUtils.setValueToXMLAttribute(messageContext, createItemElement, EWSConstants.SEND_MEETING_INVITATION,
                EWSConstants.SEND_MEETING_INVITATION_ELEMENT);
        OMElement saveItemFolderIdElement = soapFactory.createOMElement(EWSConstants.SAVE_ITEM_FOLDER_ID_ELEMENT,
                message);
        populateSaveItemFolderIdElement(messageContext, saveItemFolderIdElement);
        if (saveItemFolderIdElement.getAllAttributes().hasNext()) {
            createItemElement.addChild(saveItemFolderIdElement);
        }
        OMElement itemElement = soapFactory.createOMElement(EWSConstants.ITEMS, message);
        OMElement messageElement = soapFactory.createOMElement(EWSConstants.MESSAGE, type);
        populateMessageElement(messageContext, messageElement);
        itemElement.addChild(messageElement);
        createItemElement.addChild(itemElement);
        soapBody.addChild(createItemElement);
        return soapBody;
    }

    /**
     * used to populate message Element
     * @param messageContext message context of request
     * @param baseElement baseElement of element
     * @throws XMLStreamException
     * @throws TransformerException
     */
    private void populateMessageElement(MessageContext messageContext, OMElement baseElement) throws
            XMLStreamException, TransformerException {
        populateDirectElements(messageContext, baseElement, EWSConstants.MIME_CONTENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.ITEM_ID);
        populateDirectElements(messageContext, baseElement, EWSConstants.PARENT_FOLDER_ID);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.ITEM_CLASS, baseElement, EWSConstants
                .ITEM_CLASS_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.SUBJECT, baseElement, EWSConstants.SUBJECT_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.SENSITIVITY, baseElement, EWSConstants
                .SENSITIVITY_ELEMENT);
        OMElement bodyOmElement = soapFactory.createOMElement(EWSConstants.BODY_ELEMENT, type);
        String body = (String) ConnectorUtils.lookupTemplateParamater(messageContext, EWSConstants.BODY);
        if (!StringUtils.isEmpty(body)) {
            OMElement bodyElement = AXIOMUtil.stringToOM(body);
            String bodyType = bodyElement.getFirstChildWithName(new QName(EWSConstants.BODY_TYPE_ATTRIBUTE)).getText();
            String content = bodyElement.getFirstChildWithName(new QName(EWSConstants.CONTENT)).getText();
            bodyOmElement.addAttribute(EWSConstants.BODY_TYPE_ATTRIBUTE, bodyType, null);
            bodyOmElement.setText(content);
            baseElement.addChild(bodyOmElement);
        }
        EWSUtils.populateDirectElements(messageContext, baseElement, EWSConstants.ATTACHMENTS);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.DATE_TIME_RECEIVED, baseElement, EWSConstants
                .DATE_TIME_RECEIVED_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.SIZE, baseElement, EWSConstants.SIZE_ELEMENT);
        EWSUtils.populateDirectElements(messageContext, baseElement, EWSConstants.CATEGORIES);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IMPORTANCE, baseElement, EWSConstants
                .IMPORTANCE_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IN_REPLY_TO, baseElement, EWSConstants
                .IN_REPLY_TO_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_SUBMITTED, baseElement, EWSConstants
                .IS_SUBMITTED_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_DRAFT, baseElement, EWSConstants
                .IS_DRAFT_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_FROM_ME, baseElement, EWSConstants
                .IS_FROM_ME_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_RESEND, baseElement, EWSConstants
                .IS_RESEND_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_UNMODIFIED, baseElement, EWSConstants
                .IS_UNMODIFIED_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.INTERNET_MESSAGE_HEADERS);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.DATE_TIME_SENT, baseElement, EWSConstants
                .DATE_TIME_SENT_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.DATE_TIME_CREATED, baseElement, EWSConstants
                .DATE_TIME_CREATED_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.RESPONSE_OBJECTS);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.REMINDER_DUE_BY, baseElement, EWSConstants
                .REMINDER_DUE_BY_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.REMINDER_IS_SET, baseElement, EWSConstants
                .REMINDER_IS_SET_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.REMINDER_NEXT_TIME, baseElement, EWSConstants
                .REMINDER_NEXT_TIME_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.REMINDER_MINUTES_BEFORE_START, baseElement,
                EWSConstants.REMINDER_MINUTES_BEFORE_START_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.DISPLAY_CC, baseElement, EWSConstants
                .DISPLAY_CC_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.DISPLAY_TO, baseElement, EWSConstants
                .DISPLAY_TO_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.HAS_ATTACHMENTS, baseElement, EWSConstants
                .HAS_ATTACHMENTS_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_ASSOCIATED, baseElement, EWSConstants
                .IS_ASSOCIATED_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.WEB_CLIENT_READ_FORM_QUERY_STRING, baseElement,
                EWSConstants.WEB_CLIENT_READ_FORM_QUERY_STRING_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.WEB_CLIENT_EDIT_FORM_QUERY_STRING, baseElement,
                EWSConstants.WEB_CLIENT_EDIT_FORM_QUERY_STRING_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.CONVERSATION_ID);
        populateDirectElements(messageContext, baseElement, EWSConstants.UNIQUE_BODY);
        populateDirectElements(messageContext, baseElement, EWSConstants.FLAG);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.STORE_ENTRY_ID, baseElement, EWSConstants
                .STORE_ENTRY_ID_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.INSTANCE_KEY, baseElement, EWSConstants
                .INSTANCE_KEY_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.NORMALIZED_BODY);
        populateDirectElements(messageContext, baseElement, EWSConstants.ENTITY_EXTRACTION_RESULT);
        populateDirectElements(messageContext, baseElement, EWSConstants.POLICY_TAG);
        populateDirectElements(messageContext, baseElement, EWSConstants.ARCHIVE_TAG);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.RETENTION_DATE, baseElement, EWSConstants
                .RETENTION_DATE_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.PREVIEW, baseElement, EWSConstants.PREVIEW_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.RIGHTS_MANAGEMENT_LICENSE_DATA);
        populateDirectElements(messageContext, baseElement, EWSConstants.PREDICTED_ACTION_REASONS);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_CLUTTER, baseElement, EWSConstants
                .IS_CLUTTER_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.BLOCK_STATUS, baseElement, EWSConstants
                .BLOCK_STATUS_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.HAS_BLOCKED_IMAGES, baseElement, EWSConstants
                .HAS_BLOCKED_IMAGES_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.TEXT_BODY);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.ICON_INDEX, baseElement, EWSConstants
                .ICON_INDEX_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.SENDER);
        populateDirectElements(messageContext, baseElement, EWSConstants.TO_RECIPIENTS);
        populateDirectElements(messageContext, baseElement, EWSConstants.CC_RECIPIENTS);
        populateDirectElements(messageContext, baseElement, EWSConstants.BCC_RECIPIENTS);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_READ_RECEIPT_REQUESTED, baseElement,
                EWSConstants.IS_READ_RECEIPT_REQUESTED_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_DELIVERY_RECEIPT_REQUESTED, baseElement,
                EWSConstants.IS_DELIVERY_RECEIPT_REQUESTED_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.CONVERSATION_INDEX, baseElement, EWSConstants
                .CONVERSATION_INDEX_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.CONVERSATION_TOPIC, baseElement, EWSConstants
                .CONVERSATION_TOPIC_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.FROM);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.INTERNET_MESSAGE_ID, baseElement, EWSConstants
                .INTERNET_MESSAGE_ID_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_READ, baseElement, EWSConstants.IS_READ_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.IS_RESPONSE_REQUESTED, baseElement, EWSConstants
                .IS_RESPONSE_REQUESTED_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.REFERENCES, baseElement, EWSConstants
                .REFERENCES_ELEMENT);
        populateDirectElements(messageContext, baseElement, EWSConstants.REPLY_TO);
        populateDirectElements(messageContext, baseElement, EWSConstants.RECEIVED_BY);
        populateDirectElements(messageContext, baseElement, EWSConstants.RECEIVED_REPRESENTING);
        populateDirectElements(messageContext, baseElement, EWSConstants.APPROVAL_REQUEST_DATA);
        populateDirectElements(messageContext, baseElement, EWSConstants.VOTING_INFORMATION);
        populateDirectElements(messageContext, baseElement, EWSConstants.REMINDER_MESSAGE_DATA);


    }
}
