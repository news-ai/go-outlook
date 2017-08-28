package outlook

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"errors"
	"log"
	"net/http"

	"github.com/news-ai/tabulae/models"

	"golang.org/x/net/context"
)

type EmailAddress struct {
	EmailAddress struct {
		Address string `json:"Address"`
	} `json:"EmailAddress"`
}

type Attachment struct {
	OdataType    string `json:"@odata.type"`
	Name         string `json:"Name"`
	ContentBytes string `json:"ContentBytes"`
}

type EmailRequestError struct {
	Error struct {
		Code    string `json:"code"`
		Message string `json:"message"`
	} `json:"error"`
}

type EmailRequest struct {
	Message struct {
		Subject string `json:"Subject"`
		Body    struct {
			ContentType string `json:"ContentType"`
			Content     string `json:"Content"`
		} `json:"Body"`
		ToRecipients []EmailAddress `json:"ToRecipients"`
	} `json:"Message"`
}

type EmailRequestAttachment struct {
	Message struct {
		Subject string `json:"Subject"`
		Body    struct {
			ContentType string `json:"ContentType"`
			Content     string `json:"Content"`
		} `json:"Body"`
		ToRecipients []EmailAddress `json:"ToRecipients"`
		Attachments  []Attachment   `json:"Attachments"`
	} `json:"Message"`
}

type GetEmailRequest struct {
	OdataContext string `json:"@odata.context"`
	Value        []struct {
		OdataID   string `json:"@odata.id"`
		OdataEtag string `json:"@odata.etag"`
		ID        string `json:"Id"`
		Subject   string `json:"Subject"`
		Sender    struct {
			EmailAddress struct {
				Name    string `json:"Name"`
				Address string `json:"Address"`
			} `json:"EmailAddress"`
		} `json:"Sender"`
		OdataType string `json:"@odata.type,omitempty"`
	} `json:"value"`
}

func (o *Outlook) SendEmail(c context.Context, from string, to string, subject string, body string, email models.Email) error {
	if len(o.AccessToken) > 0 {
		toEmail := EmailAddress{}
		toEmail.EmailAddress.Address = to

		var message EmailRequest

		message.Message.Subject = subject
		message.Message.ToRecipients = append(message.Message.ToRecipients, toEmail)

		message.Message.Body.ContentType = "HTML"
		message.Message.Body.Content = body

		messageJson, err := json.Marshal(message)
		if err != nil {
			log.Printf("%v", err)
			return err
		}

		messageQuery := bytes.NewReader(messageJson)

		URL := BASEURL + "api/v2.0/me/sendmail"
		req, _ := http.NewRequest("POST", URL, messageQuery)

		req.Header.Add("Authorization", "Bearer "+o.AccessToken)
		req.Header.Add("Content-Type", "application/json")

		client := http.Client{}
		response, err := client.Do(req)
		if err != nil {
			log.Printf("%v", err)
			return err
		}
		defer response.Body.Close()

		if response.StatusCode == 200 || response.StatusCode == 202 {
			log.Printf("%v", response.Body)
			return nil
		}

		// Decode JSON from Google
		decoder := json.NewDecoder(response.Body)
		var emailRequestError EmailRequestError
		err = decoder.Decode(&emailRequestError)
		if err != nil {
			log.Printf("%v", err)
			return err
		}

		log.Printf("%v", emailRequestError.Error)

		return errors.New("Email could not be sent")
	}

	return errors.New("No access token supplied")
}

func (o *Outlook) SendEmailWithAttachments(c context.Context, from string, to string, subject string, body string, email models.Email, files []models.File, bytesArray [][]byte, attachmentType []string, fileNames []string) error {
	if len(o.AccessToken) > 0 {
		toEmail := EmailAddress{}
		toEmail.EmailAddress.Address = to

		attachments := []Attachment{}
		if len(files) > 0 {
			for x := 0; x < len(bytesArray); x++ {
				str := base64.StdEncoding.EncodeToString(bytesArray[x])

				singleAttachment := Attachment{}
				singleAttachment.Name = fileNames[x]
				singleAttachment.OdataType = "#Microsoft.OutlookServices.FileAttachment"
				singleAttachment.ContentBytes = str

				attachments = append(attachments, singleAttachment)
			}
		}

		var message EmailRequestAttachment

		message.Message.Subject = subject
		message.Message.ToRecipients = append(message.Message.ToRecipients, toEmail)

		message.Message.Body.ContentType = "HTML"
		message.Message.Body.Content = body

		message.Message.Attachments = attachments

		messageJson, err := json.Marshal(message)
		if err != nil {
			log.Printf("%v", err)
			return err
		}

		messageQuery := bytes.NewReader(messageJson)

		URL := BASEURL + "api/v2.0/me/sendmail"
		req, _ := http.NewRequest("POST", URL, messageQuery)

		req.Header.Add("Authorization", "Bearer "+o.AccessToken)
		req.Header.Add("Content-Type", "application/json")

		client := http.Client{}
		response, err := client.Do(req)
		if err != nil {
			log.Printf("%v", err)
			return err
		}
		defer response.Body.Close()

		if response.StatusCode == 200 || response.StatusCode == 202 {
			return nil
		}

		// Decode JSON from Google
		decoder := json.NewDecoder(response.Body)
		var emailRequestError EmailRequestError
		err = decoder.Decode(&emailRequestError)
		if err != nil {
			log.Printf("%v", err)
			return err
		}

		log.Printf("%v", emailRequestError.Error)

		return errors.New("Email could not be sent")
	}

	return errors.New("No access token supplied")
}

func (o *Outlook) GetEmail(c context.Context, to string, subject string) error {
	if len(o.AccessToken) > 0 {
		URL := BASEURL + "api/v2.0/me/MailFolders/sentitems/messages/?$select=Sender,Subject&$search=\"subject:" + subject + "\""
		req, _ := http.NewRequest("GET", URL, nil)

		req.Header.Add("Authorization", "Bearer "+o.AccessToken)
		req.Header.Add("Content-Type", "application/json")

		client := http.Client{}
		response, err := client.Do(req)
		if err != nil {
			log.Printf("%v", err)
			return err
		}
		defer response.Body.Close()
	}
	return nil
}
