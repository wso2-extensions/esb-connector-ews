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

import javax.xml.stream.XMLStreamException;
import javax.xml.transform.TransformerException;

/**
 * Class used to generate GetAttachment Operation Soap Request
 */
public class GetAttachmentMediator extends AbstractConnector {
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
     * @throws TransformerException when transformation couldn't be done
     */
    private SOAPBody populateBody(MessageContext messageContext) throws XMLStreamException, TransformerException {
        SOAPBody soapBody = soapFactory.createSOAPBody();
        OMElement getAttachmentOmElement = soapFactory.createOMElement(EWSConstants.GET_ATTACHMENTS, message);
        OMElement attachmentShapeOmElement = soapFactory.createOMElement(EWSConstants.ATTACHMENT_SHAPE, message);
        populateAttachmentShape(messageContext, attachmentShapeOmElement);
        getAttachmentOmElement.addChild(attachmentShapeOmElement);
        OMElement attachmentIdsOmElement = soapFactory.createOMElement(EWSConstants.ATTACHMENT_IDS_ELEMENT, message);
        OMElement attachmentIdOmElement = soapFactory.createOMElement(EWSConstants.ATTACHMENT_ID_ELEMENT, type);
        EWSUtils.setValueToXMLAttribute(messageContext, attachmentIdOmElement, EWSConstants.ATTACHMENT_ID_,
                EWSConstants.ID_ATTRIBUTE);
        attachmentIdsOmElement.addChild(attachmentIdOmElement);
        getAttachmentOmElement.addChild(attachmentIdsOmElement);
        soapBody.addChild(getAttachmentOmElement);
        return soapBody;
    }

    /**
     * Used to populate AttachmentScope Element
     * @param messageContext message context of request
     * @param baseElement baseElement of attachmentShape
     * @throws XMLStreamException
     * @throws TransformerException when transformation couldn't be done
     */
    private void populateAttachmentShape(MessageContext messageContext, OMElement baseElement) throws
            XMLStreamException, TransformerException {
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.INCLUDE_MIME_CONTENT, baseElement, EWSConstants
                .INCLUDE_MIME_CONTENT_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.BODY_TYPE, baseElement, EWSConstants
                .BODY_TYPE_ELEMENT);
        EWSUtils.setValueToXMLElement(messageContext, EWSConstants.FILTER_HTML_CONTENT, baseElement, EWSConstants
                .FILTER_HTML_CONTENT);
        String additionalProperties = (String) ConnectorUtils.lookupTemplateParamater(messageContext, EWSConstants
                .ADDITIONAL_PROPERTIES);
        if (!StringUtils.isEmpty(additionalProperties)) {
            OMElement additionalPropertiesElement = AXIOMUtil.stringToOM(additionalProperties);
            baseElement.addChild(EWSUtils.setNameSpaceForElements(additionalPropertiesElement));
        }
    }
}
