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
import org.apache.axiom.soap.SOAPBody;
import org.apache.axiom.soap.SOAPEnvelope;
import org.apache.axiom.soap.SOAPFactory;
import org.apache.axiom.soap.SOAPHeader;
import org.apache.axis2.AxisFault;
import org.apache.synapse.MessageContext;
import org.wso2.carbon.connector.core.AbstractConnector;
import org.wso2.carbon.connector.core.ConnectException;

import javax.xml.stream.XMLStreamException;
import javax.xml.transform.TransformerException;

import static org.wso2.carbon.connector.EWSUtils.populateItemIds;
import static org.wso2.carbon.connector.EWSUtils.populateItemShape;

/**
 * This class used to generate GetItem Operation SOAP request
 */
public class GetItemMediator extends AbstractConnector {
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
     * @throws TransformerException throws when
     */
    private SOAPHeader populateSoapHeader(MessageContext messageContext) throws XMLStreamException,
            TransformerException {
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
     * Used to populate soap Body
     * @param messageContext message context of request
     * @return Soap Body
     * @throws XMLStreamException
     * @throws TransformerException throws when
     */
    private SOAPBody populateBody(MessageContext messageContext) throws XMLStreamException, TransformerException {
        SOAPBody soapBody = soapFactory.createSOAPBody();
        OMElement getItemElement = soapFactory.createOMElement(EWSConstants.GET_ITEM_ELEMENT, message);
        getItemElement.addChild(populateItemShape(messageContext));
        getItemElement.addChild(populateItemIds(messageContext));
        soapBody.addChild(getItemElement);
        return soapBody;
    }


}
