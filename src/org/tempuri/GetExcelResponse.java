
package org.tempuri;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlElement;
import jakarta.xml.bind.annotation.XmlRootElement;
import jakarta.xml.bind.annotation.XmlType;


/**
 * <p>anonymous complex type 的 Java 類別.
 * 
 * <p>下列綱要片段會指定此類別中包含的預期內容.
 * 
 * <pre>
 * &lt;complexType&gt;
 *   &lt;complexContent&gt;
 *     &lt;restriction base="{http://www.w3.org/2001/XMLSchema}anyType"&gt;
 *       &lt;sequence&gt;
 *         &lt;element name="GetExcelResult" type="{http://tempuri.org/}PSSApiResponse" minOccurs="0"/&gt;
 *       &lt;/sequence&gt;
 *     &lt;/restriction&gt;
 *   &lt;/complexContent&gt;
 * &lt;/complexType&gt;
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "", propOrder = {
    "getExcelResult"
})
@XmlRootElement(name = "GetExcelResponse")
public class GetExcelResponse {

    @XmlElement(name = "GetExcelResult")
    protected PSSApiResponse getExcelResult;

    /**
     * 取得 getExcelResult 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link PSSApiResponse }
     *     
     */
    public PSSApiResponse getGetExcelResult() {
        return getExcelResult;
    }

    /**
     * 設定 getExcelResult 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link PSSApiResponse }
     *     
     */
    public void setGetExcelResult(PSSApiResponse value) {
        this.getExcelResult = value;
    }

}
