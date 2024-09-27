import datetime
import sys

from asn1crypto import cms
from dateutil.parser import parse
from pypdf import PdfReader


class AttrClass:
    def __init__(self, data, cls_name=None):
        self._data = data
        self._cls_name = cls_name

    def __getattr__(self, name):
        try:
            value = self._data[name]
        except KeyError:
            value = None
        else:
            if isinstance(value, dict):
                return AttrClass(value, cls_name=name.capitalize() or self._cls_name)
        return value

    def __values_for_str__(self):
        """Values to show for "str" and "repr" methods"""
        return [
            (k, v) for k, v in self._data.items()
            if isinstance(v, (str, int, datetime.datetime))
        ]

    def __str__(self):
        """String representation of object"""
        values = ", ".join([
            f"{k}={v}" for k, v in self.__values_for_str__()
        ])
        return f"{self._cls_name or self.__class__.__name__}({values})"

    def __repr__(self):
        return f"<{self}>"


class Signature(AttrClass):
    """Signature helper class

    Attributes:
        type (str): 'timestamp' or 'signature'
        signing_time (datetime, datetime): when user has signed
            (user HW's clock)
        signer_name (str): the signer's common name
        signer_contact_info (str, None): the signer's email / contact info
        signer_location (str, None): the signer's location
        signature_type (str): ETSI.cades.detached, adbe.pkcs7.detached, ...
        certificate (Certificate): the signers certificate
        digest_algorithm (str): the digest algorithm used
        message_digest (bytes): the digest
        signature_algorithm (str): the signature algorithm used
        signature_bytes (bytest): the raw signature
    """

    @property
    def signer_name(self):
        return (
            self._data.get('signer_name') or
            getattr(self.certificate.subject, 'common_name', '')
        )


class Subject(AttrClass):
    """Certificate subject helper class

    Attributes:
        common_name (str): the subject's common name
        given_name (str): the subject's first name
        surname (str): the subject's surname
        serial_number (str): subject's identifier (may not exist)
        country (str): subject's country
    """
    pass


class Certificate(AttrClass):
    """Signer's certificate helper class

    Attributes:
        version (str): v3 (= X509v3)
        serial_number (int): the certificate's serial number
        subject (object): signer's subject details
        issuer (object): certificate issuer's details
        signature (object): certificate signature
        extensions (list[OrderedDict]): certificate extensions
        validity (object): validity (not_before, not_after)
        subject_public_key_info (object): public key info
        issuer_unique_id (object, None): issuer unique id
        subject_uniqiue_id (object, None): subject unique id
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.subject = Subject(self._data['subject'])

    def __values_for_str__(self):
        return (
            super().__values_for_str__() +
            [('common_name', self.subject.common_name)]
        )
        
class SignatureDetails:
    def __init__(self):
        self.digest_algorithm = None
        self.signature_algorithm = None
        self.content_type = None
        self.type = None
        self.signer_contact_info = None
        self.signer_location = None
        self.signing_time = None
        self.signature_type = None
        self.signature_handler = None
        self.issuer_country_name = None
        self.issuer_organization_name = None
        self.issuer_common_name = None
        self.subject_country_name = None
        self.subject_organization_name = None
        self.subject_organizational_unit_name = None
        self.subject_common_name = None
        self.subject_locality_name = None
        self.common_name = None
        self.valid_from = None
        self.valid_to = None

    def to_dict(self):
        return {
            "digest_algorithm": self.digest_algorithm,
            "signature_algorithm": self.signature_algorithm,
            "content_type": self.content_type,
            "type": self.type,
            "signer_contact_info": self.signer_contact_info,
            "signer_location": self.signer_location,
            "signing_time": self.signing_time,
            "signature_type": self.signature_type,
            "signature_handler": self.signature_handler,
            "valid_from": self.valid_from,
            "valid_to": self.valid_to,
            "issuer": {
                "country_name": self.issuer_country_name,
                "organization_name": self.issuer_organization_name,
                "common_name": self.issuer_common_name,
            },
            "subject": {
                "country_name": self.subject_country_name,
                "organization_name": self.subject_organization_name,
                "organizational_unit_name": self.subject_organizational_unit_name,
                "common_name": self.subject_common_name,
                "locality_name": self.subject_locality_name,
            },
        }

    def __repr__(self):
        return (f"SignatureDetails(digest_algorithm={self.digest_algorithm}, signature_algorithm={self.signature_algorithm}, "
                f"content_type={self.content_type}, type={self.type}, signer_contact_info={self.signer_contact_info}, "
                f"signer_location={self.signer_location}, signing_time={self.signing_time}, signature_type={self.signature_type}, "
                f"signature_handler={self.signature_handler}, issuer_country_name={self.issuer_country_name}, "
                f"issuer_organization_name={self.issuer_organization_name}, issuer_common_name={self.issuer_common_name}, "
                f"subject_country_name={self.subject_country_name}, subject_organization_name={self.subject_organization_name}, "
                f"subject_organizational_unit_name={self.subject_organizational_unit_name}, subject_common_name={self.subject_common_name}, "
                f"subject_locality_name={self.subject_locality_name})")

    
    
class SignatureExtract:
    def __init__(self) -> None:
        pass
        
    def parse_pkcs7_signatures(self, signature_data: bytes):
        content_info = cms.ContentInfo.load(signature_data).native
        if content_info['content_type'] != 'signed_data':
            return None
        content = content_info['content']
        certificates = content['certificates']
        signer_infos = content['signer_infos']
        for signer_info in signer_infos:
            sid = signer_info['sid']
            digest_algorithm = signer_info['digest_algorithm']['algorithm']
            signature_algorithm = signer_info['signature_algorithm']['algorithm']
            signature_bytes = signer_info['signature']
            signed_attrs = {
                sa['type']: sa['values'][0] for sa in signer_info['signed_attrs']}
            for cert in certificates:
                cert = cert['tbs_certificate']
                if (
                    sid['serial_number'] == cert['serial_number'] and
                    sid['issuer'] == cert['issuer']
                ):
                    break
            else:
                raise RuntimeError(
                    f"Couldn't find certificate in certificates collection: {sid}")
            yield dict(
                sid=sid,
                certificate=Certificate(cert),
                digest_algorithm=digest_algorithm,
                signature_algorithm=signature_algorithm,
                signature_bytes=signature_bytes,
                signer_info=signer_info,
                **signed_attrs,
            )


    def get_pdf_signatures(self, filename):
        """Parse PDF signatures"""
        reader = PdfReader(filename)
        fields = reader.get_fields().values()
        signature_field_values = [
            f.value for f in fields if f.field_type == '/Sig']
        for v in signature_field_values:
            v_type = v['/Type']
            if v_type in ('/Sig', '/DocTimeStamp'):  # unknow types are skipped
                is_timestamp = v_type == '/DocTimeStamp'
                try:
                    signing_time = parse(v['/M'][2:].strip("'").replace("'", ":"))
                except KeyError:
                    signing_time = None
                # - used standard for signature encoding, in my case:
                # - get PKCS7/CMS/CADES signature package encoded in ASN.1 / DER format
                raw_signature_data = v['/Contents']
                # if is_timestamp:
                for attrdict in self.parse_pkcs7_signatures(raw_signature_data):
                    if attrdict:
                        attrdict.update(dict(
                            type='timestamp' if is_timestamp else 'signature',
                            signer_name=v.get('/Name'),
                            signer_contact_info=v.get('/ContactInfo'),
                            signer_location=v.get('/Location'),
                            signing_time=signing_time or attrdict.get('signing_time'),
                            signature_type=v['/SubFilter'][1:],  # ETSI.CAdES.detached, ...
                            signature_handler=v['/Filter'][1:],
                            raw=raw_signature_data,
                        ))
                        yield Signature(attrdict)
    
    
    
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <filename>")
        sys.exit(1)
    filename = sys.argv[1]
    signatureExtract = SignatureExtract()
    for signature in signatureExtract.get_pdf_signatures(filename):
        print(f"--- {signature.type} ---")
        print(f"Signature: {signature}")
        print(f"Signer: {signature.signer_name}")
        print(f"Signing time: {signature.signing_time}")
        certificate = signature.certificate
        print(f"Signer's certificate: {certificate}")
        print(f"  - not before: {certificate.validity.not_before}")
        print(f"  - not after: {certificate.validity.not_after}")
        print(f"  - issuer: {certificate.issuer}")
        subject = signature.certificate.subject
        print(f"  - subject: {subject}")
        print(f"    - common name: {subject.common_name}")
        print(f"    - serial number: {subject.serial_number}")