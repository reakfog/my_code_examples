# Django model
class Order(TimestampedModel):
    """
    Client create this, when he wants to buy something.
    """

    STATUSES = Choices(
        ("DRAFT", _("Draft")),
        ("RESERVED", _("Reserved")),
        ("CONFIRMED", _("Confirmed")),
        ("PAID", _("Paid")),
        ("IN_PROGRESS", _("In delivery")),
        ("CANCELED", _("Canceled")),
        ("DELETED", _("Deleted")),
    )

    STOCK_RESERVED_STATUSES = (STATUSES.CONFIRMED, STATUSES.IN_PROGRESS, STATUSES.PAID)
    EDITABLE_STATUSES = (STATUSES.DRAFT, STATUSES.RESERVED)

    organization = ForeignKey(Organization, CASCADE, related_name="orders")
    manager = ForeignKey(User, CASCADE, related_name="orders", null=True)
    offers = ManyToManyField(Offer, through="OrderItem", through_fields=("order", "offer"), related_name="orders")
    documents = ManyToManyField(Document, related_name="orders")

    title = CharField(max_length=255, null=True, blank=True)
    comment = TextField(null=True)

    status = FSMField(_("Статус"), max_length=15, choices=STATUSES, default=STATUSES.DRAFT, protected=True)
    status_updated_at = MonitorField(monitor="status", blank=True, null=True)
    dereservation_at = DateTimeField(null=True, default=None, editable=False)

    transactions = GenericRelation(BatchTransaction, related_query_name="order")

    objects = OrderManager()

    class Meta:
        db_table = "product_order"
        verbose_name = _("order")
        verbose_name_plural = _("orders")

    def __str__(self) -> str:
        return f"{self.id} {self.organization_id} {self.status}"

    def save(self, *args: Any, **kwargs: Any) -> None:
        if self.id:
            # Добавляем status_updated_at в update_fields, если они указаны без status_updated_at
            update_fields: Union[list, tuple, set] = kwargs.get("update_fields")
            if update_fields and "status_updated_at" not in update_fields:
                if not isinstance(update_fields, set):
                    update_fields = set(update_fields)
                kwargs["update_fields"] = update_fields.union({"status_updated_at"})
        super().save(*args, **kwargs)

    @property
    def is_confirmed(self) -> bool:
        return self.status == self.STATUSES.CONFIRMED

    @property
    def _is_cancelable(self) -> bool:
        return not self.deliveries.filter(status__in=(Delivery.STATUSES.FACT, Delivery.STATUSES.SHIPPED)).exists()

    @staticmethod
    def get_new_dereservation_time() -> datetime:
        return timezone.now() + timedelta(hours=1)

    @transition(status, STATUSES.DRAFT, STATUSES.RESERVED)
    def mark_reserved(self) -> None:
        self.items.filter(amount=0).delete()
        self.dereservation_at = self.get_new_dereservation_time()
        self.items.update(status=OrderItem.STATUSES.RESERVE)

    @transition(status, STATUSES.RESERVED, STATUSES.DRAFT)
    def mark_draft(self) -> None:
        self.dereservation_at = None
        self.items.update(status=OrderItem.STATUSES.NEW)

    @transition(status, (STATUSES.DRAFT, STATUSES.RESERVED), STATUSES.CONFIRMED)
    def make_confirmed(self) -> None:
        from foodex.apps.external.tasks import add_deal, ext_dbs_sync_order
        from foodex.apps.logistic.utils import make_batch_items

        self.items.filter(amount=0).delete()
        for order_item in self.items.all():
            order_item.update_price_and_vat()
            order_item.save(update_fields=["price", "vat"])

    @transition(status, STATUSES.CONFIRMED, STATUSES.PAID)
    def mark_paid(self) -> None:
        from foodex.apps.external.tasks import add_delivery_paid

        deliveries = self.deliveries.exclude(status=Delivery.STATUSES.CANCELED)
        deliveries.update(status=Delivery.STATUSES.PAID)

    @transition(status, (STATUSES.PAID, STATUSES.CONFIRMED, STATUSES.IN_PROGRESS), STATUSES.CANCELED)
    def mark_canceled(self) -> None:
        if not self._is_cancelable:
            raise TransitionNotAllowed("Order is not cancelable.")

        self.transactions.all().delete()

        for delivery in self.deliveries.all():
            delivery.mark_canceled()
            delivery.save(update_fields=["status"])

    @transition(status, (STATUSES.DRAFT, STATUSES.RESERVED), STATUSES.DELETED)
    def mark_deleted(self) -> None:
        self.deliveries.all().update(status=Delivery.STATUSES.DELETED)

 
class OrderItem(UpdateFieldsMixin, FieldSetterMixin, TimestampedModel):
    """
    Unit of order
    """

    STATUSES = Choices(
        ("NEW", _("New")),
        ("ACCEPT", _("Accept")),
        ("REJECT", _("Reject")),
        ("RESERVE", _("Reserve")),
    )

    FIELDS_TO_SET_DEFAULT = ("price", "vat", "package_id", "sum", "vat_sum")

    offer = ForeignKey(Offer, CASCADE, related_name="order_items")
    order = ForeignKey(Order, CASCADE, related_name="items")
    package = ForeignKey(ProductPackage, SET_NULL, related_name="order_items", null=True)

    amount = DecimalField(_("Количество"), max_digits=14, decimal_places=3)
    price = DecimalField(_("Цена"), max_digits=17, decimal_places=2)
    vat = DecimalField(_("НДС"), max_digits=5, decimal_places=2)

    sum = DecimalField(_("Сумма"), max_digits=28, decimal_places=2)
    vat_sum = DecimalField(_("Сумма НДС"), max_digits=30, decimal_places=4)

    status = CharField(choices=STATUSES, default=STATUSES.NEW, max_length=16, null=True)

    _is_price_changed = False
    _is_vat_changed = False

    class Meta:
        db_table = "product_order_item"
        verbose_name = _("order item")
        verbose_name_plural = _("order items")

    def __str__(self) -> str:
        return "{}({}) {}".format(self.offer, self.offer.product, self.amount)

    def save(self, *args: Any, **kwargs: Any) -> None:
        if self.order.status == Order.STATUSES.RESERVED:
            if not self.id:
                self.status = self.STATUSES.RESERVE

            if self.order.dereservation_at is not None:
                self.order.dereservation_at = self.order.get_new_dereservation_time()
                self.order.save(update_fields=["dereservation_at"])

        if self.is_price_changed:
            self.set_sum()
            self.add_update_fields.add("sum")
            if self.is_vat_changed:
                self.set_vat_sum()
                self.add_update_fields.add("vat_sum")

        super().save(*args, **kwargs)


# Django viewset
class OrderViewSet(CreateModelMixin, RetrieveModelMixin, UpdateModelMixin, DestroyModelMixin, GenericViewSet):
    """
    Order endpoint
    """

    queryset = (
        Order.objects.annotate(sum=Sum("items__sum"))
        .select_related("manager")
        .prefetch_related(
            "documents",
            "items",
            Prefetch("deliveries", Delivery.objects.only("id", "order_id").all()),
        )
        .all()
    )
    serializer_class = OrderPageOrderSerializer
    deliveries_serializer_class = OrderPageDeliverySerializer

    permission_classes = (IsInObjectOwnerOrganizationOrHasPermission,)
    pagination_class = BaseSetPagination

    def get_queryset(self) -> QuerySet:
        if self.action == "generate_deliveries":
            return Order.objects.filter(status=Order.STATUSES.DRAFT)
        return super().get_queryset()

    def get_serializer_class(self, *args: Any, **kwargs: Any) -> type[ModelSerializer]:
        if self.action == "generate_deliveries":
            return self.deliveries_serializer_class
        return self.serializer_class

    @transaction.atomic
    def destroy(self, request: Request, *args: Any, **kwargs: Any) -> Response:
        order = self.get_object()
        try:
            order.mark_deleted()
        except TransitionNotAllowed as e:
            raise ValidationError(_("Order cannot be deleted.")) from e

        order.save(update_fields=["status"])
        logger_info.info("Order has been deleted.", extra={"order_id": order.id})
        return Response(status=HTTP_204_NO_CONTENT)

    @extend_schema(request=None)
    @action(["post"], detail=True)
    @base_action
    @transaction.atomic
    def cancel(self, instance: Order, request: Request, *args: Any, **kwargs: Any) -> None:
        try:
            instance.mark_canceled()
        except TransitionNotAllowed as e:
            raise ValidationError(_("Order cannot be canceled.")) from e
        logger_info.info("Order has been canceled.", extra={"order_id": instance.id})
        instance.save(update_fields=["status"])

    @extend_schema(request=None)
    @action(["post"], detail=True)
    def export_draft_order(self, request: Request, *args: Any, **kwargs: Any) -> HttpResponse:
        """
        Экспорт данных о черновике заказа в формат excel
        """
        order = self.get_object()
        file = FreshLogicExporter.export_draft_order(order)

        response = HttpResponse(file, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response["Content-Disposition"] = f"attachment; filename=order_{order.id}.xlsx"
        return response

    @extend_schema(request=None)
    @action(["post"], detail=True)
    @base_action
    @transaction.atomic
    def mark_reserved(self, instance: Order, *args: Any, **kwargs: Any) -> None:
        """
        Order status transfer from DRAFT to RESERVED.
        """
        instance.mark_reserved()
        instance.save(update_fields=["status", "dereservation_at"])

    @extend_schema(request=None)
    @action(["post"], detail=True)
    @base_action
    @transaction.atomic
    def confirm(self, instance: Order, *args: Any, **kwargs: Any) -> None:
        """
        Order confirmation
        """
        instance.make_confirmed()
        instance.save(update_fields=["status"])


# Djanfo form
class BatchForm(forms.ModelForm):
    product = forms.ModelChoiceField(
        required=True,
        queryset=Product.objects.order_by("title"),
        widget=forms.Select(attrs={"class": "form-control select2"}),
    )
    queryset = Organization.objects.filter(status=Organization.STATUSES.ACTIVE).order_by("title")
    organization = forms.ModelChoiceField(
        required=True,
        queryset=queryset,
        widget=forms.Select(attrs={"class": "form-control select2"}),
    )
    storage = forms.ModelChoiceField(
        required=True,
        queryset=Storage.objects.all().order_by("title"),
        widget=forms.Select(attrs={"class": "form-control select2"}),
    )
    amount = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={"placeholder": _("Amount"), "class": "form-control"}),
    )
    estimate_price = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={"placeholder": _("Estimated price"), "class": "form-control"}),
    )
    product_created_at = forms.DateTimeField(
        required=False,
        widget=forms.DateTimeInput(
            format="%Y-%m-%d %H:%M:%S",
            attrs={
                "placeholder": _("Product created at"),
                "class": "form-control",
                "autocomplete": "off",
                "type": "date",
            },
        ),
    )
    product_expired_at = forms.DateTimeField(
        required=False,
        widget=forms.DateTimeInput(
            format="%Y-%m-%d %H:%M:%S",
            attrs={
                "placeholder": _("Product expired at"),
                "class": "form-control",
                "autocomplete": "off",
                "type": "date",
            },
        ),
    )
    gtd_code = forms.CharField(
        required=False, widget=forms.TextInput(attrs={"placeholder": _("GTD code"), "class": "form-control"})
    )
    status = forms.ChoiceField(
        required=True,
        choices=Batch.STATUSES,
        widget=forms.Select(attrs={"class": "form-control select2"}),
    )

    def __init__(self, *args: t.Any, **kwargs: t.Any):
        self.user = kwargs.pop("user", None)
        super().__init__(*args, **kwargs)
        user_timezone = self.user.timezone if self.user else None
        self.fields["gtd_code"].validators = [
            GtdCodeRegexValidator(
                regex=Batch.GTD_CODE_REGEX, message=_("GTD code is incorrect."), user_timezone=user_timezone
            )
        ]

    def clean_gtd_code(self) -> t.Optional[str]:
        return self.cleaned_data.get("gtd_code") or None

    def clean(self) -> dict:
        cleaned_data = super().clean()
        product = cleaned_data["product"]
        gtd_code = cleaned_data.get("gtd_code")

        if not product.country.is_gtd_required and gtd_code is not None:
            raise forms.ValidationError(_("GTD code field is only for imported products."))
        elif product.country.is_gtd_required and gtd_code is None:
            raise forms.ValidationError(_("GTD code field is required."))

        return cleaned_data

    class Meta:
        model = Batch
        fields = [
            "product",
            "organization",
            "storage",
            "amount",
            "estimate_price",
            "product_created_at",
            "product_expired_at",
            "gtd_code",
            "status",
        ]
  
 
# Store data exporter
class FreshLogicExporter:
    PAGE_SIZE = 44
    ENDING_ROW_PADDING = 6
    TABLE_ROW = 14

    BASE_DRAFT_ORDER_PATH = os.path.join(settings.ASSETS_ROOT, "external/base_draft_order.xlsx")
    BASE_ORDER_PATH = os.path.join(settings.ASSETS_ROOT, "external/base_order.xlsx")

    ORDER_DATE_TEXT = "Заказ №{} от {}"
    ORDER_SUPPLIER_TEXT = 'Поставщик: ООО "Фудекс" ИНН 9703021089'
    ORDER_CLIENT_TEXT = "Покупатель: {}"
    ORDER_TOTAL_WO_VAT_TITLE = "Итого:"
    ORDER_TOTAL_WO_VAT_TEXT = "{0:.2f} руб."
    ORDER_TOTAL_VAT_TITLE = "В том числе НДС:"
    ORDER_TOTAL_VAT_TEXT = "{0:.2f} руб."
    ORDER_TOTAL_TITLE = "Всего к оплате:"
    ORDER_TOTAL_TEXT = "{0:.2f} руб."

    ORDER_END_TEXT = (
        "Данное предложение актуально на момент выгрузки. "
        "Резервирование товара происходит в момент заказа и действует до оплаты счета."
    )

    STOCKS_TITLE = "Сверка остатков платформы FX и WMS на {}"

    def __init__(self, warehouse: "FreshLogicWarehouse") -> None:
        self.warehouse = warehouse
    
    @classmethod
    def export_draft_order(cls, order: Order) -> bytes:
        xw = XlsxWriter(filename=cls.BASE_DRAFT_ORDER_PATH)

        xw.write_cell(
            cls.ORDER_DATE_TEXT.format(order.id, order.created_at.strftime("%d.%m.%Y %H:%M")), column=1, row=7
        )
        hyperlink = f"{settings.BASE_URL}/dashboard/orders/{order.id}/creation"
        xw.write_cell(cls.ORDER_SUPPLIER_TEXT, column=1, row=9)
        xw.write_cell(hyperlink, column=4, row=9, font_size=10, hyperlink=hyperlink)
        xw.write_cell(cls.ORDER_CLIENT_TEXT.format(order.organization.title), column=1, row=10)

        xw.next_print_area(from_row=1, size=cls.PAGE_SIZE)

        row = cls.TABLE_ROW
        order_items = OrderItem.objects.select_related("offer").filter(order=order)
        order_total_price = 0
        order_total_vat = 0
        for i, order_item in enumerate(order_items, start=1):
            product = order_item.offer.product
            total = order_item.offer.price_for_transport_package * order_item.amount
            total_vat = Decimal(total / (100 + product.vat)).quantize(Decimal(".01")) * product.vat
            order_total_price += total
            order_total_vat += total_vat

            xw.write_cell(i, column=1, row=row, align="center", font_size=9, bordered=True)
            xw.write_cell(product.code, column=2, row=row, align="center", font_size=9, bordered=True)
            xw.write_cell(product.title, column=3, row=row, align="left", font_size=10, bordered=True)
            xw.write_cell(float(order_item.amount), column=4, row=row, align="center", font_size=9, bordered=True)
            xw.write_cell(
                float(order_item.offer.price_for_transport_package),
                column=5,
                row=row,
                align="center",
                font_size=9,
                bordered=True,
            )
            xw.write_cell(
                float(total),
                column=6,
                row=row,
                align="center",
                font_size=9,
                bordered=True,
            )
            xw.write_cell(
                localtime(order_item.offer.expired_at).strftime("%d.%m.%Y"),
                column=7,
                row=row,
                align="center",
                font_size=9,
                bordered=True,
            )

            row += 1
            if row % cls.PAGE_SIZE == 0:
                xw.next_print_area(from_row=row + 1, size=cls.PAGE_SIZE)
                row += 2

        after_items = row
        order_sum = float(Decimal(order_total_price).quantize(Decimal(".001")))
        order_vat = float(Decimal(order_total_vat).quantize(Decimal(".001")))

        total = divmod(order_sum, 1)
        total_vat = divmod(order_vat, 1)

        xw.merge_cells(start_row=(after_items + 1), start_column=4, end_column=5)
        xw.merge_cells(start_row=(after_items + 1), start_column=6, end_column=7)
        xw.write_cell(cls.ORDER_TOTAL_WO_VAT_TITLE, column=4, row=(after_items + 1), align="right", font_size=10)
        xw.write_cell(
            cls.ORDER_TOTAL_WO_VAT_TEXT.format(order_sum),
            column=6,
            row=(after_items + 1),
            align="left",
            font_size=10,
        )

        xw.merge_cells(start_row=(after_items + 2), start_column=4, end_column=5)
        xw.merge_cells(start_row=(after_items + 2), start_column=6, end_column=7)
        xw.write_cell(cls.ORDER_TOTAL_VAT_TITLE, column=4, row=(after_items + 2), align="right", font_size=10)
        xw.write_cell(
            cls.ORDER_TOTAL_VAT_TEXT.format(order_vat),
            column=6,
            row=(after_items + 2),
            align="left",
            font_size=10,
        )

        xw.merge_cells(start_row=(after_items + 3), start_column=4, end_column=5)
        xw.merge_cells(start_row=(after_items + 3), start_column=6, end_column=7)
        xw.write_cell(cls.ORDER_TOTAL_TITLE, column=4, row=(after_items + 3), align="right", font_size=10)
        xw.write_cell(
            cls.ORDER_TOTAL_TEXT.format(order_sum),
            column=6,
            row=(after_items + 3),
            align="left",
            font_size=10,
        )

        last_row = after_items + cls.ENDING_ROW_PADDING
        xw.write_cell(cls.ORDER_END_TEXT, column=2, row=last_row, font_size=10)

        current_page = math.ceil(last_row / cls.PAGE_SIZE)
        if current_page > xw.page_num:
            xw.next_print_area(from_row=after_items + 2, size=cls.PAGE_SIZE)

        return xw.get_buffer()
