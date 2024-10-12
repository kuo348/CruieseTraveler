
package org.tempuri;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
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
 *         &lt;element name="sDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="eDate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="cruMode" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="shipType" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="cruCate" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="vesType" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="firstCru" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="dirRoute" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="vesselCname" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="vesselNo" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="callSign" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
 *         &lt;element name="port" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/&gt;
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
    "sDate",
    "eDate",
    "cruMode",
    "shipType",
    "cruCate",
    "vesType",
    "firstCru",
    "dirRoute",
    "vesselCname",
    "vesselNo",
    "callSign",
    "port"
})
@XmlRootElement(name = "GetExcel")
public class GetExcel {

    protected String sDate;
    protected String eDate;
    protected String cruMode;
    protected String shipType;
    protected String cruCate;
    protected String vesType;
    protected String firstCru;
    protected String dirRoute;
    protected String vesselCname;
    protected String vesselNo;
    protected String callSign;
    protected String port;

    /**
     * 取得 sDate 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getSDate() {
        return sDate;
    }

    /**
     * 設定 sDate 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setSDate(String value) {
        this.sDate = value;
    }

    /**
     * 取得 eDate 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getEDate() {
        return eDate;
    }

    /**
     * 設定 eDate 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setEDate(String value) {
        this.eDate = value;
    }

    /**
     * 取得 cruMode 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getCruMode() {
        return cruMode;
    }

    /**
     * 設定 cruMode 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setCruMode(String value) {
        this.cruMode = value;
    }

    /**
     * 取得 shipType 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getShipType() {
        return shipType;
    }

    /**
     * 設定 shipType 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setShipType(String value) {
        this.shipType = value;
    }

    /**
     * 取得 cruCate 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getCruCate() {
        return cruCate;
    }

    /**
     * 設定 cruCate 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setCruCate(String value) {
        this.cruCate = value;
    }

    /**
     * 取得 vesType 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getVesType() {
        return vesType;
    }

    /**
     * 設定 vesType 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setVesType(String value) {
        this.vesType = value;
    }

    /**
     * 取得 firstCru 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getFirstCru() {
        return firstCru;
    }

    /**
     * 設定 firstCru 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setFirstCru(String value) {
        this.firstCru = value;
    }

    /**
     * 取得 dirRoute 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getDirRoute() {
        return dirRoute;
    }

    /**
     * 設定 dirRoute 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setDirRoute(String value) {
        this.dirRoute = value;
    }

    /**
     * 取得 vesselCname 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getVesselCname() {
        return vesselCname;
    }

    /**
     * 設定 vesselCname 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setVesselCname(String value) {
        this.vesselCname = value;
    }

    /**
     * 取得 vesselNo 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getVesselNo() {
        return vesselNo;
    }

    /**
     * 設定 vesselNo 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setVesselNo(String value) {
        this.vesselNo = value;
    }

    /**
     * 取得 callSign 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getCallSign() {
        return callSign;
    }

    /**
     * 設定 callSign 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setCallSign(String value) {
        this.callSign = value;
    }

    /**
     * 取得 port 特性的值.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    public String getPort() {
        return port;
    }

    /**
     * 設定 port 特性的值.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    public void setPort(String value) {
        this.port = value;
    }

}
