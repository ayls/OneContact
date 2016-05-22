var currentInputValue = ""
var textBox;
var list;
var token;
var timeoutHandle;

function bootstrap() {
	textBox = $("textarea.cp_primaryInput");
	textBox.on("change keyup paste", function () { handleInputChange(this); });
}

function handleInputChange(that) {
	var newInputValue = $(that).val();
	if(newInputValue == currentInputValue) {
    	return;
	}

	currentInputValue = newInputValue;

	if (timeoutHandle)
		window.clearTimeout(timeoutHandle);

	timeoutHandle = window.setTimeout(updateList, 500);
}

function updateList() {
	if (!token)
		return;

	$.ajax({
         url: "https://outlook.office.com/api/v2.0/me/contacts?$select=EmailAddresses,GivenName,Surname&$filter=GivenName%20eq%20'" + currentInputValue + "'%20or%20Surname%20eq%20'" + currentInputValue + "'",
         type: "GET",
         beforeSend: function (request) {
            request.setRequestHeader("Accept", "application/json");
            request.setRequestHeader("Authorization", "Bearer " + token);
         },
         success: handleResponse
      });
}

function handleResponse(data) {
	list = $("ul[role='listbox']");
	if (!list)
		return;

	var isListCleared = false;
	if (data.value.length > 0) {
		var idx = 0;
		for (i = 0; i < data.value.length; i++) {
			var contact = data.value[i];
			for (j = 0; j < contact.EmailAddresses.length; j++) {
				if (contact.EmailAddresses[j].Address) {
					if (!isListCleared) {
						$(list).empty();
						isListCleared = true;
					}
					addToList(idx++, j, contact);
				}
			}
		}
	}
}

function addToList(idx, addressIdx, contact) {
	var listId = $(list).attr('id');
	var contactId = contact.Id;
	var dataId = contactId + "_" + addressIdx;
	var itemId = listId + dataId;
	var detailId = "icTmCPV2_contactList" + idx;
	var detailTitleId = detailId + "_usertilecontainer";
	var detailTitleLinkId = detailId + "_frame_clip";
	var detailTitleImgId = detailId + "_usertile";
	var detailTitleBarId = detailId + "_bar";
	var contactFullName = contact.GivenName + ((contact.Surname != null) ? (' ' + contact.Surname) : '');

	var el = $('<li id="' + itemId + '" role="option" class="CL_Row" style="top:0px; left:0px" data-unique-id="' + dataId + '">'
		+ '<div class="CL_Contact">'    
			+ '<table class="CL_Contact_Tbl" cellspacing="0" cellpadding="0" unselectable="on" aria-label="' + contact.GivenName + '">'        
				+ '<colgroup>'
					+ '<col class="CL_User_Tile_Col">'
					+ '<col class="CL_Display_Name_Col">'
					+ '<col class="CL_Remove_Col">'
				+ '</colgroup>'        
				+ '<tbody>'            
					+ '<tr>'                
						+ '<td>'
							+ '<div class="CL_User_Tile">'
								+ '<div id="' + detailId + '" class="c_ic c_ic_h_ml icpwc">'
									+ '<div class="c_ic_img_h c_ic_img_ml">'
										+ '<div class="c_ic_img_sub c_ic_img_ml" id="' + detailTitleId + '">'
											+ '<a id="' + detailTitleLinkId + '" title="" class="c_ic_frame_clip c_ml" target="_top" href="#" onclick="return false;">'
												+ '<div class="c_ic_tile_clip" style="cursor: pointer;">'
													+ '<img id="' + detailTitleImgId + '" errsrc="https://a.gfx.ms/ic/bluemanm.png" alt="" src="https://a.gfx.ms/ic/bluemanm.png" class="c_ic_tile">'
												+ '</div>'
												+ '<div class="c_ic_blueframe c_ic_bar" id="' + detailTitleBarId + '" dir="ltr" style="visibility: inherit;"></div>'
											+ '</a>'
										+ '</div>'
									+ '</div>'
									+ '<div class="c_clr"></div>'
								+ '</div>'
							+ '</div>'
						+ '</td>'                
						+ '<td>'                    
							+ '<div class="CL_Display_Name t_atc">' + contactFullName + '</div>'                                         
							+ '<div class="CL_Email t_atc">' + contact.EmailAddresses[addressIdx].Address + '</div>'                                    
						+ '</td>'                                    
						+ '<td>'
							+ '<a class="CL_Remove">'
								+ '<div class="cl_ImgStrip">&nbsp;</div>'
							+ '</a>'
						+ '</td>'                            
					+ '</tr>'        
				+ '</tbody>'    
			+ '</table>'
		+ '</div>'
	+ '</li>');

	el.on("click", function() {
    	textBox.val(contact.EmailAddresses[addressIdx].Address + ";");
		currentinputvalue = "";
	});

	el.appendTo(list);
}

$( document ).ready(function() {
	bootstrap()
});

chrome.runtime.onMessage.addListener(function (msg, sender) {
	token = msg;
});