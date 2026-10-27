// Outlook POC: operates only on the current item after the user clicks a button.
(() => {
  const address = 'testws@scb.co.th';
  const status = document.getElementById('outlook-action-status');
  const subjectField = document.getElementById('outlook-subject');
  const bccButton = document.getElementById('outlook-add-bcc');
  const subjectButton = document.getElementById('outlook-get-subject');
  let ready = false;
  const report = (message, error = false) => {
    status.textContent = message;
    status.style.color = error ? '#b42318' : '';
  };
  function item() {
    if (!ready || !window.Office?.context?.mailbox?.item) {
      report('เปิด add-in จากอีเมลใน Outlook ก่อนใช้งาน', true);
      return null;
    }
    return Office.context.mailbox.item;
  }
  function failed(result) { return result.status !== Office.AsyncResultStatus.Succeeded; }
  function onReady() {
    if (!window.Office?.onReady) return;
    Office.onReady(info => {
      ready = info.host === Office.HostType.Outlook;
      if (ready) {
        document.getElementById('outlook-poc-actions')?.removeAttribute('hidden');
        report('พร้อมใช้งานกับอีเมลปัจจุบัน');
      }
    });
  }
  window.addEventListener('office-script-loaded', onReady);
  if (window.Office?.onReady) onReady();

  subjectButton.addEventListener('click', () => {
    const current = item();
    if (!current) return;
    if (typeof current.subject === 'string') {
      subjectField.value = current.subject;
      report('ดึงหัวข้ออีเมลแล้ว');
    } else if (current.subject?.getAsync) {
      current.subject.getAsync(result => {
        if (failed(result)) return report(`ดึงหัวข้อไม่สำเร็จ: ${result.error?.message || 'เกิดข้อผิดพลาด'}`, true);
        subjectField.value = result.value || '';
        report('ดึงหัวข้ออีเมลแล้ว');
      });
    } else report('ไม่สามารถอ่านหัวข้อจากอีเมลนี้ได้', true);
  });

  bccButton.addEventListener('click', () => {
    const current = item();
    if (!current) return;
    if (!current.bcc?.getAsync || !current.bcc?.addAsync) {
      report('ปุ่ม Bcc ใช้ได้ขณะเขียนหรือตอบอีเมลเท่านั้น', true);
      return;
    }
    bccButton.disabled = true;
    current.bcc.getAsync(result => {
      if (failed(result)) {
        bccButton.disabled = false;
        return report(`อ่าน Bcc ไม่สำเร็จ: ${result.error?.message || 'เกิดข้อผิดพลาด'}`, true);
      }
      if ((result.value || []).some(recipient => recipient.emailAddress?.toLowerCase() === address)) {
        bccButton.disabled = false;
        return report(`${address} อยู่ใน Bcc แล้ว`);
      }
      current.bcc.addAsync([address], added => {
        bccButton.disabled = false;
        if (failed(added)) return report(`เพิ่ม Bcc ไม่สำเร็จ: ${added.error?.message || 'เกิดข้อผิดพลาด'}`, true);
        report(`เพิ่ม ${address} ใน Bcc แล้ว`);
      });
    });
  });
})();
