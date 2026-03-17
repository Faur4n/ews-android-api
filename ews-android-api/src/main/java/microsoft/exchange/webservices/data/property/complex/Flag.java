package microsoft.exchange.webservices.data.property.complex;

import java.util.Date;

import microsoft.exchange.webservices.data.core.EwsServiceXmlReader;
import microsoft.exchange.webservices.data.core.EwsServiceXmlWriter;
import microsoft.exchange.webservices.data.core.EwsUtilities;
import microsoft.exchange.webservices.data.core.XmlElementNames;
import microsoft.exchange.webservices.data.core.enumeration.misc.XmlNamespace;
import microsoft.exchange.webservices.data.core.enumeration.property.ItemFlagStatus;

/**
 * Represents the Flag on a message.
 */
public final class Flag extends ComplexProperty {

    /**
     * The flag status.
     */
    private ItemFlagStatus flagStatus;

    private Date startDate;

    private Date dueDate;

    private Date completeDate;

    /**
     * Initializes a new instance.
     */
    public Flag() {
    }


    /**
     * Tries to read element from XML.
     *
     * @param reader the reader
     * @return true, if successful
     * @throws Exception the exception
     */
    @Override
    public boolean tryReadElementFromXml(EwsServiceXmlReader reader)
            throws Exception {
        if (reader.getLocalName().equals(XmlElementNames.FlagStatus)) {
            this.flagStatus = reader.readElementValue(ItemFlagStatus.class);
            return true;
        }
        if (reader.getLocalName().equals(XmlElementNames.StartDate)) {
            this.startDate = reader.readElementValueAsDateTime();
            return true;
        }
        if (reader.getLocalName().equals(XmlElementNames.DueDate)) {
            this.dueDate = reader.readElementValueAsDateTime();
            return true;
        }
        if (reader.getLocalName().equals(XmlElementNames.CompleteDate)) {
            this.completeDate = reader.readElementValueAsDateTime();
            return true;
        }
        return false;
    }

    @Override
    public void writeElementsToXml(EwsServiceXmlWriter writer)
            throws Exception {
        writer.writeElementValue(XmlNamespace.Types, XmlElementNames.FlagStatus, flagStatus);

        if (flagStatus == ItemFlagStatus.Flagged) {
            writer.writeElementValue(XmlNamespace.Types, XmlElementNames.StartDate, startDate);
            writer.writeElementValue(XmlNamespace.Types, XmlElementNames.DueDate, dueDate);
        }else if(flagStatus == ItemFlagStatus.Complete) {
            writer.writeElementValue(XmlNamespace.Types, XmlElementNames.CompleteDate, completeDate);
        }
    }

    @Override
    protected void internalValidate() throws Exception {
        super.internalValidate();
        EwsUtilities.validateParam(this.flagStatus, "FlagStatus");
    }

    public ItemFlagStatus getFlagStatus() {
        return this.flagStatus;
    }

    public void setFlagStatus(ItemFlagStatus flagStatus) {
        if (this.canSetFieldValue(this.flagStatus, flagStatus)){
            this.flagStatus = flagStatus;
            this.changed();
        }
    }

    public Date getStartDate() {
        return startDate;
    }

    public void setStartDate(Date startDate) {
        if (this.canSetFieldValue(this.startDate, startDate)) {
            this.startDate = startDate;
            this.changed();
        }
    }

    public Date getCompleteDate() {
        return completeDate;
    }

    public void setCompleteDate(Date completeDate) {
        if (this.canSetFieldValue(this.completeDate, completeDate)) {
            this.completeDate = completeDate;
            this.changed();
        }
    }

    public Date getDueDate() {
        return dueDate;
    }

    public void setDueDate(Date dueDate) {
        if (this.canSetFieldValue(this.dueDate, dueDate)) {
            this.dueDate = dueDate;
            this.changed();
        }
    }
}