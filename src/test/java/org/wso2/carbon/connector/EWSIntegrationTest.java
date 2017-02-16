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

import org.apache.axiom.om.OMElement;
import org.apache.axiom.om.util.AXIOMUtil;
import org.apache.commons.io.IOUtils;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.DefaultHttpClient;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import org.wso2.connector.integration.test.base.ConnectorIntegrationTestBase;
import org.wso2.connector.integration.test.base.RestResponse;

import javax.xml.namespace.QName;
import java.io.File;
import java.io.FileInputStream;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.Map;

/**
 * Sample integration test
 */
public class EWSIntegrationTest extends ConnectorIntegrationTestBase {

    private Map<String, String> esbRequestHeadersMap = new HashMap<String, String>();
    private Map<String, String> namespaceMap = new HashMap<String, String>();
    private String itemId, changeKey;

    @BeforeClass(alwaysRun = true)
    public void setEnvironment() throws Exception {
        init("ews-connector-1.0.0");
        esbRequestHeadersMap.put("Accept-Charset", "UTF-8");
        esbRequestHeadersMap.put("Content-Type", "text/xml");
        namespaceMap.put("m", "microsoft.com/exchange/services/2006/messages");
        namespaceMap.put("t", "http://schemas.microsoft.com/exchange/services/2006/types");
    }

    @Test(enabled = true, groups = {"wso2.esb"}, description = "EWS test case")
    public void testCreateItem() throws Exception {
        DefaultHttpClient defaultHttpClient = new DefaultHttpClient();
        String createItemProxyUrl = getProxyServiceURL("createItemOperation");
        RestResponse<OMElement> esbSoapResponse =
                sendXmlRestRequest(createItemProxyUrl, "POST", esbRequestHeadersMap, "CreateItem.xml");
        OMElement omElement = esbSoapResponse.getBody();
        OMElement createItemResponseMessageOmelement = omElement.getFirstElement().getFirstChildWithName(new QName
                ("http://schemas.microsoft.com/exchange/services/2006/messages", "CreateItemResponseMessage", "m"));
        String success = createItemResponseMessageOmelement.getAttributeValue(new QName("ResponseClass"));
        Assert.assertEquals(success, "Success");
        itemId = (String) xPathEvaluate(omElement, "string(//t:ItemId/@Id)", namespaceMap);
        changeKey = (String) xPathEvaluate(omElement, "string(//t:ItemId/@ChangeKey)", namespaceMap);
        Assert.assertNotNull(itemId);
        Assert.assertNotNull(changeKey);
        String createAttachmentOperation = getProxyServiceURL("createAttachmentOperation");
        FileInputStream fileInputStream = new FileInputStream(getESBResourceLocation() + File.separator + "config"
                + File.separator + "restRequests" + File.separator + "sampleRequest" + File.separator
                + "CreateAttachment.xml");
        OMElement createAttachmentSoapRequest = AXIOMUtil.stringToOM(IOUtils.toString(fileInputStream));
        OMElement parentItemID = createAttachmentSoapRequest.getFirstChildWithName(new QName("ParentItemId"));
        parentItemID.getFirstChildWithName(new QName("Id")).setText(itemId);
        parentItemID.getFirstChildWithName(new QName("ChangeKey")).setText(changeKey);
        HttpPost httpPost = new HttpPost(createAttachmentOperation);
        httpPost.setEntity(new StringEntity(createAttachmentSoapRequest.toString(), ContentType.TEXT_XML.withCharset
                (Charset.forName("UTF-8"))));
        HttpResponse response = defaultHttpClient.execute(httpPost);
        OMElement createAttachmentResponse = AXIOMUtil.stringToOM(IOUtils.toString(response.getEntity().getContent()));
        String createAttachmentStatus = createAttachmentResponse.getFirstElement().getFirstElement()
                .getAttributeValue(new QName("ResponseClass"));
        String attachmentId = (String) xPathEvaluate(createAttachmentResponse, "string(//t:AttachmentId/@Id)",
                namespaceMap);
        Assert.assertEquals(createAttachmentStatus, "Success");
        String getAttachmentProxyUrl = getProxyServiceURL("getAttachment");
        String payload = "<GetAttachment><AttachmentId>" + attachmentId + "</AttachmentId></GetAttachment>";
        HttpPost getAttachmentPost = new HttpPost(getAttachmentProxyUrl);
        getAttachmentPost.setEntity(new StringEntity(payload, ContentType.TEXT_XML.withCharset(Charset.forName
                ("UTF-8"))));
        HttpResponse getAttachmentPostHttpResponse = defaultHttpClient.execute(getAttachmentPost);
        OMElement getAttachmentOm = AXIOMUtil.stringToOM(IOUtils.toString(getAttachmentPostHttpResponse.getEntity()
                .getContent()));
        String GetAttachmentStatus = getAttachmentOm.getFirstElement().getFirstElement().getAttributeValue(new QName("ResponseClass"));
        Assert.assertEquals(GetAttachmentStatus, "Success");
    }

    @Test(enabled = true, groups = {"wso2.esb"}, description = "EWS test case", dependsOnMethods = {"testCreateItem"})
    public void testFindItemAndSendItem() throws Exception {
        DefaultHttpClient defaultHttpClient = new DefaultHttpClient();
        String createItemProxyUrl = getProxyServiceURL("findItemOperation");
        RestResponse<OMElement> esbSoapResponse = sendXmlRestRequest(createItemProxyUrl, "POST",
                esbRequestHeadersMap, "FindItem.xml");
        OMElement omElement = esbSoapResponse.getBody();
        String findItemStatus = omElement.getFirstElement().getFirstElement().getAttributeValue(new QName("ResponseClass"));
        Assert.assertEquals(findItemStatus, "Success");
        itemId = (String) xPathEvaluate(omElement, "string(//t:ItemId/@Id)", namespaceMap);
        changeKey = (String) xPathEvaluate(omElement, "string(//t:ItemId/@ChangeKey)", namespaceMap);
        Assert.assertNotNull(itemId);
        Assert.assertNotNull(changeKey);
        String sendItemOperation = getProxyServiceURL("sendItem");
        HttpPost httpPost = new HttpPost(sendItemOperation);
        String payload = "<SendItem><SaveItemToFolder>true</SaveItemToFolder><ItemId><Id>" + itemId +
                "</Id><ChangeKey>"+changeKey+"</ChangeKey></ItemId></SendItem>";
        httpPost.setEntity(new StringEntity(payload, ContentType.TEXT_XML.withCharset(Charset.forName("UTF-8"))));
        HttpResponse response = defaultHttpClient.execute(httpPost);
        OMElement createAttachmentResponse = AXIOMUtil.stringToOM(IOUtils.toString(response.getEntity().getContent()));
        String GetAttachmentStatus = createAttachmentResponse.getFirstElement().getFirstElement().getAttributeValue(new QName("ResponseClass"));
        Assert.assertEquals(GetAttachmentStatus, "Success");
    }
}