/* Version: 16.0.15307.10000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/


Office._appMode = {
    Read: 1,
    Compose: 2,
    ReadCompose: 3
}

Office._cast_item = function () {
    this.toAppointmentCompose = function (item) {
        return new Office._$MailboxAppointment(Office._appMode.Compose);
    }
    this.toAppointmentRead = function (item) {
        return new Office._$MailboxAppointment(Office._appMode.Read);
    }
    this.toAppointment = function (item) {
        return new Office._$MailboxAppointment(Office._appMode.ReadCompose);
    }
    this.toMessageCompose = function (item) {
        return new Office._$MailboxMessage(Office._appMode.Compose);
    }
    this.toMessageRead = function (item) {
        return new Office._$MailboxMessage(Office._appMode.Read);
    }
    this.toMessage = function (item) {
        return new Office._$MailboxMessage(Office._appMode.ReadCompose);
    }
    this.toItemCompose = function (item) {
        return new Office._$MailboxItem(Office._appMode.Compose);
    }
    this.toItemRead = function (item) {
        return new Office._$MailboxItem(Office._appMode.Read);
    }
};

Office._context_mailbox_item = function () {
    Office._$MailboxItem_helper(this, Office._appMode.ReadCompose);
    Office._$MailboxAppointment_helper(this, Office._appMode.ReadCompose);
    Office._$MailboxMessage_helper(this, Office._appMode.ReadCompose);
};

Office._$MailboxItem = function (appMode) {
    Office._$MailboxItem_helper(this, appMode);
    Office._$MailboxAppointment_helper(this, appMode);
    Office._$MailboxMessage_helper(this, appMode);
}

Office._$MailboxAppointment = function (appMode) {
    Office._$MailboxItem_helper(this, appMode);
    Office._$MailboxAppointment_helper(this, appMode);
}

Office._$MailboxMessage = function (appMode) {
    Office._$MailboxItem_helper(this, appMode);
    Office._$MailboxMessage_helper(this, appMode);
}

Office._$MailboxItem_helper = function (obj, appMode) {
    // Field documentation ------------------------------------------

    // Attachments property.
    attachmentsDoc = {
        attachments_read: {
            conditions: {
                hosts: ["outlook; not outlookcompose"]
            },
            name: "attachments",
            annotate: {
                ///<field name="attachments" type='AttachmentDetails[]'>Gets a list of attachments to the item.</field>
                attachments: undefined
            },
            contents: function () {
                return new Array(new Office._context_mailbox_item_attachmentDetails())
            }
        },
        attachments_read_compose: {
            conditions: {
                hosts: ["outlook; outlookcompose"]
            },
            name: "attachments",
            annotate: {
                ///<field name="attachments">Gets a list of attachments to the item. In compose mode the attachments property is undefined. In read mode it returns a list of attachments to the item.</field>
                attachments: undefined
            },
            contents: function () {
                return new Array(new Office._context_mailbox_item_attachmentDetails())
            }
        }
    }

    bodyDoc = {
        body_compose: {
            conditions: {
                hosts: ["not outlook, outlookcompose"]
            },
            name: "body",
            annotate: {
                /// <field name="body" type='Body'>Provides methods to get and set the body of the item.</field>
                body: undefined
            },
            contents: function () {
                return new Office._context_mailbox_item_body()
            }
        },
        body_read_compose: {
            conditions: {
                hosts: ["outlook, outlookcompose"]
            },
            name: "body",
            annotate: {
                /// <field name="body"> Gets the content of an item. In read mode, the body property is undefined. In compose mode it returns a Body object.</field>
                body: undefined
            },
            contents: function () {
                return new Office._context_mailbox_item_body()
            }
        }
    }

    // dateTimeCreated property.
    dateTimeCreatedDoc = {
        dateTimeCreated_read: {
            conditions: {
                hosts: ["outlook, not outlookcompose"]
            },
            name: "dateTimeCreated",
            annotate: {
                ///<field name="dateTimeCreated" type='Date'>Gets the date and time that the item was created.</field>
                dateTimeCreated: undefined
            },
            contents: function () {
                return new Date()
            }
        },
        dateTimeCreated_read_compose: {
            conditions: {
                hosts: ["outlook, outlookcompose"]
            },
            name: "dateTimeCreated",
            annotate: {
                ///<field name="dateTimeCreated" type='Date'>Gets the date and time that the item was created. In compose mode the dateTimeCreated property is undefined.</field>
                dateTimeCreated: undefined
            },
            contents: function () {
                return new Date()
            }
        }
    }

    // dateTimeModified property.
    dateTimeModifiedDoc = {
        dateTimeModified_read: {
            conditions: {
                hosts: ["outlook, not outlookcompose"]
            },
            name: "dateTimeModified",
            annotate: {
                ///<field name="dateTimeModified" type='Date'>Gets the date and time that the item was last modified.</field>
                dateTimeModified: undefined
            },
            contents: function () {
                return new Date()
            }
        },
        dateTimeModified_read_compose: {
            conditions: {
                hosts: ["outlook, outlookcompose"]
            },
            name: "dateTimeModified",
            annotate: {
                ///<field name="dateTimeModified" type='Date'>Gets the date and time that the item was last modified. In compose mode the dateTimeModified property is undefined.</field></field>
                dateTimeModified: undefined
            },
            contents: function () {
                return new Date()
            }
        }
    }

    // itemClass property.
    itemClassDoc = {
        itemClass_read: {
            conditions: {
                hosts: ["outlook, not outlookcompose"]
            },
            name: "itemClass",
            annotate: {
                ///<field name="itemClass" type='String'>Gets the item class of the item.</field>
                itemClass: undefined
            }
        },
        itemClass_read_compose: {
            conditions: {
                hosts: ["outlook, outlookcompose"]
            },
            name: "itemClass",
            annotate: {
                ///<field name="itemClass" type='String'>Gets the item class of the item. In compose mode the itemClass property is undefined.</field>
                itemClass: undefined
            }
        }
    }

    // itemId property
    // This property is in all modes, so it gets processed in place rather than
    // after the extra documentation is removed.
    Office._processContents(obj, {
        itemIdDoc: {
            conditions: {
                hosts: ["outlook", "outlookcompose"]
            },
            name: "itemId",
            annotate: {
                ///<field name="itemId" type='String'>Gets the Exchange Web Services (EWS) item identifier of an item.</field>
                itemId: undefined
            }
        }
    })

    // itemType property.
    Office._processContents(obj, {
        itemTypeDoc: {
            conditions: {
                hosts: ["outlook", "outlookcompose"]
            },
            name: "itemType",
            annotate: {
                ///<field name="itemType" type='Office.MailboxEnums.ItemType'>Gets the type of an item that an instance represents.</field>
                itemType: undefined
            }
        }
    })

    //obj.normalizedSubject = {};
    normalizedSubjectDoc = {
        normalizedSubject_read: {
            conditions: {
                hosts: ["outlook, not outlookcompose"]
            },
            name: "normalizedSubject",
            annotate: {
                ///<field name="normalizedSubject" type='String'>Gets the subject of the item, with standard prefixes removed.</field>
                normalizedSubject: undefined
            }
        },
        normalizedSubject_read_compose: {
            conditions: {
                hosts: ["outlook, outlookcompose"]
            },
            name: "normalizedSubject",
            annotate: {
                ///<field name="normalizedSubject" type='String'>Gets the subject of the item, with standard prefixes removed. In compose mode, the normalizedSubject property is undefined.</field>
                normalizedSubject: undefined
            }
        }
    }

    subjectDoc = {
        subject_compose: {
            conditions: {
                hosts: ["not outlook, outlookcompose"]
            },
            name: "subject",
            annotate: {
                /// <field name="subject" type='Subject'>Provides methods to get and set the item subject.</field>
                subject: undefined
            },
            contents: function () {
                return new Office._context_mailbox_item_subject()
            }
        },
        subject_read: {
            conditions: {
                hosts: ["outlook, not outlookcompose"]
            },
            name: "subject",
            annotate: {
                /// <field name="subject" type='String'>Gets the subject of the item.</field>
                subject: undefined
            }
        },
        subject_read_compose: {
            conditions: {
                hosts: ["outlook, outlookcompose"]
            },
            name: "subject",
            annotate: {
                /// <field name="subject"> Gets the subject of an item. In compose mode the Subject property returns a Subject object. In read mode, it returns a string.</field>
                subject: undefined
            }
        }
    }

    if (appMode == Office._appMode.Compose) {
        delete attachmentsDoc["attachments_read"];
        delete attachmentsDoc["attachments_read_compose"];
        delete bodyDoc.body_compose["conditions"];
        delete bodyDoc["body_read_compose"];
        delete dateTimeCreatedDoc["dateTimeCreated_read"];
        delete dateTimeCreatedDoc["dateTimeCreated_read_compose"];
        delete dateTimeModifiedDoc["dateTimeModified_read"];
        delete dateTimeModifiedDoc["dateTimeModified_read_compose"];
        delete itemClassDoc["itemClass_read"];
        delete itemClassDoc["itemClass_read_compose"];
        delete normalizedSubjectDoc["normalizedSubject_read"];
        delete normalizedSubjectDoc["normalizedSubject_read_compose"];
        delete subjectDoc.subject_compose["conditions"];
        delete subjectDoc["subject_read"];
        delete subjectDoc["subject_read_compose"];
    }
    else if (appMode == Office._appMode.Read) {
        delete attachmentsDoc.attachments_read["conditions"];
        delete attachmentsDoc["attachments_read_compose"];
        delete bodyDoc["body_compose"];
        delete bodyDoc["body_read_compose"];
        delete dateTimeCreatedDoc.dateTimeCreated_read["conditions"];
        delete dateTimeCreatedDoc["dateTimeCreated_read_compose"];
        delete dateTimeModifiedDoc.dateTimeModified_read["conditions"];
        delete dateTimeModifiedDoc["dateTimeModified_read_compose"];
        delete itemClassDoc.itemClass_read["conditions"];
        delete itemClassDoc["itemClass_read_compose"];
        delete normalizedSubjectDoc.normalizedSubject_read["conditions"];
        delete normalizedSubjectDoc["normalizedSubject_read_compose"];
        delete subjectDoc.subject_read["conditions"];
        delete subjectDoc["subject_compose"];
        delete subjectDoc["subject_read_compose"];
    }
    else if (appMode == Office._appMode.ReadCompose) {
        delete attachmentsDoc["attachments_read"];
        delete attachmentsDoc.attachments_read_compose["conditions"];
        delete bodyDoc["body_compose"];
        delete bodyDoc.body_read_compose["conditions"];
        delete dateTimeCreatedDoc["dateTimeCreated_read"];
        delete dateTimeCreatedDoc.dateTimeCreated_read_compose["conditions"];
        delete dateTimeModifiedDoc["dateTimeModified_read"];
        delete dateTimeModifiedDoc.dateTimeModified_read_compose["conditions"];
        delete itemClassDoc["itemClass_read"];
        delete itemClassDoc.itemClass_read_compose["conditions"];
        delete normalizedSubjectDoc["normalizedSubject_read"];
        delete normalizedSubjectDoc.normalizedSubject_read_compose["conditions"];
        delete subjectDoc["subject_compose"];
        delete subjectDoc["subject_read"];
        delete subjectDoc.subject_read_compose["conditions"];
    }

    Office._processContents(obj, attachmentsDoc);
    Office._processContents(obj, bodyDoc);
    Office._processContents(obj, dateTimeCreatedDoc);
    Office._processContents(obj, dateTimeModifiedDoc);
    Office._processContents(obj, itemClassDoc);
    Office._processContents(obj, normalizedSubjectDoc);
    Office._processContents(obj, subjectDoc);

    if (appMode == Office._appMode.Compose || appMode == Office._appMode.ReadCompose) {
        obj.addFileAttachmentAsync = function (uri, attachmentName, options, callback) {
            ///<summary>Attach a file to an item.</summary>
            ///<param name="uri" type="String">A URI that provides the location of the file. Required.</param>
            ///<param name="attachmentName" type="String">The name to display while the attachment is loading. The name is limited to 256 characters. Required.</param>
            ///<param name="options" type="Object" optional="true">An optional parameters or state data passed to the callback method. Optional.</param>
            ///<param name="callback" type="function" optional="true">The method to invoke when the attachment finishes uploading. Optional.</param>

            var result = new Office._Mailbox_AsyncResult("attachmentAsync");
            if (arguments.length == 3) { callback = options; }
            callback(result);
        };

        obj.addItemAttachmentAsync = function (itemId, attachmentName, options, callback) {
            ///<summary>Attach an email item to an item.</summary>
            ///<param name="itemId" type="string">The Exchange identifier of the item to attach. The maximum length is 100 characters.</param>
            ///<param name="attachmentName" type="string">The name to display while the attachment is loading. The name is limited to 256 characters. </param>
            ///<param name="options" type="Object" optional="true">An optional parameters or state data passed to the callback method. </param>
            ///<param name="callback" type="function" optional="true">The method to invoke when the attachment finishes uploading. </param>

            var result = new Office._Mailbox_AsyncResult("attachmentAsync");
            if (arguments.length == 3) { callback = options; }
            callback(result);
        };

        obj.removeAttachmentAsync = function (attachmentIndex, options, callback) {
            ///<summary>Removes a file or item that was previously atta