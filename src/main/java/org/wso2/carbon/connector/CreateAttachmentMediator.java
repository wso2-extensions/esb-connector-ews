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

/**
 * Class used to Create CreateAttachment Soap Request
 */
public class CreateAttachmentMediator extends AbstractConnector {
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
     * @throws TransformerException
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
     * Used to populate soap Body
     * @param messageContext message context of request
     * @return Soap Body
     * @throws XMLStreamException
     * @throws TransformerException
     */
    private SOAPBody populateBody(MessageContext messageContext) throws XMLStreamException, TransformerException {
        SOAPBody soapBody = soapFactory.createSOAPBody();
        OMElement createAttachment = soapFactory.createOMElement(EWSConstants.CREATE_ATTACHMENT_ELEMENT, message);
        OMElement parentItemIdElement = soapFactory.createOMElement(EWSConstants.PARENT_ITEM_ID_ELEMENT, message);
        String parentItemId = (String) ConnectorUtils.lookupTemplateParamater(messageContext, EWSConstants
                .PARENT_ITEM_ID);
        if (!StringUtils.isEmpty(parentItemId)) {
            OMElement parentIdOmElement = AXIOMUtil.stringToOM(parentItemId);
            String id = parentIdOmElement.getFirstChildWithName(new QName(EWSConstants.ID_ATTRIBUTE)).getText();
            String changeKey = parentIdOmElement.getFirstChildWithName(new QName(EWSConstants.CHANGE_KEY_ATTRIBUTE))
                    .getText();
            parentItemIdElement.addAttribute(soapFactory.createOMAttribute(EWSConstants.ID_ATTRIBUTE, null, id));
            parentItemIdElement.addAttribute(soapFactory.createOMAttribute(EWSConstants.CHANGE_KEY_ATTRIBUTE, null,
                    changeKey));
        }
        createAttachment.addChild(parentItemIdElement);
        EWSUtils.populateDirectElements(messageContext, createAttachment, EWSConstants.ATTACHMENTS, message);
        soapBody.addChild(createAttachment);
        return soapBody;
    }

}
