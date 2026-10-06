<?php
namespace GDMexico\SepomexImport\Controller\Adminhtml\Import;
use Magento\Backend\App\Action;
use Magento\Framework\View\Result\PageFactory;
class Index extends Action
{
    const ADMIN_RESOURCE = 'GDMexico_SepomexImport::import';
    private $pageFactory;
    public function __construct(Action\Context $context, PageFactory $pageFactory) { parent::__construct($context); $this->pageFactory = $pageFactory; }
    public function execute() { $page = $this->pageFactory->create(); $page->setActiveMenu('GDMexico_SepomexImport::import'); $page->getConfig()->getTitle()->prepend(__('Importar SEPOMEX')); return $page; }
}
