<?php
namespace GDMexico\SepomexImport\Controller\Adminhtml\Import;
use Magento\Backend\App\Action;
use Magento\Framework\App\Action\HttpPostActionInterface;
use GDMexico\SepomexImport\Model\Importer;
class Execute extends Action implements HttpPostActionInterface
{
    const ADMIN_RESOURCE = 'GDMexico_SepomexImport::import';
    private $importer;
    public function __construct(Action\Context $context, Importer $importer) { parent::__construct($context); $this->importer = $importer; }
    public function execute()
    {
        try {
            $file = $this->getRequest()->getFiles('sepomex_file');
            if (!$file || empty($file['tmp_name'])) { throw new \RuntimeException('Seleccione un archivo XLS o XLSX.'); }
            $result = $this->importer->execute($file);
            $this->messageManager->addSuccessMessage(__('SEPOMEX actualizado: %1 registros, %2 CP. Respaldo: sepomex_backup.', $result['records'], $result['cps']));
        } catch (\Throwable $e) {
            $this->messageManager->addErrorMessage(__('No se importó SEPOMEX: %1', $e->getMessage()));
        }
        return $this->resultRedirectFactory->create()->setPath('sepomex/import/index');
    }
}
