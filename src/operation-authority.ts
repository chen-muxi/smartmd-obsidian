export interface RunState {
    runAllowed?: boolean;
    confirmed?: boolean;
}

interface RunControl {
    generation: number;
    finalTickets: Set<string>;
    stopWaiters: Set<() => void>;
}

export interface RunTicket {
    state: RunState;
    generation: number;
    serviceVersion: number;
    contextKey: string;
    id: string;
}

export class OperationAuthority {
    private serviceAllowed = true;
    private serviceVersion = 0;
    private controls = new WeakMap<RunState, RunControl>();
    private knownControls = new Set<RunControl>();

    canRun(state: RunState | null | undefined): boolean {
        return this.serviceAllowed && state?.runAllowed === true && state?.confirmed === true;
    }

    start(state: RunState, contextKey = ""): RunTicket {
        if (!this.canRun(state)) throw new Error("当前项目尚未获得运行许可或人工确认");
        const control = this.control(state);
        return {
            state,
            generation: control.generation,
            serviceVersion: this.serviceVersion,
            contextKey,
            id: `${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`,
        };
    }

    isCurrent(ticket: RunTicket): boolean {
        return this.canRun(ticket.state)
            && this.control(ticket.state).generation === ticket.generation
            && this.serviceVersion === ticket.serviceVersion;
    }

    require(ticket: RunTicket): void {
        if (!this.isCurrent(ticket)) throw new Error("运行许可已撤销，正式落地已停止");
    }

    confirm(state: RunState): void {
        const control = this.control(state);
        control.generation++;
        state.confirmed = true;
        state.runAllowed = true;
    }

    async pause(state: RunState): Promise<void> {
        await this.stop(state);
    }

    async stop(state: RunState): Promise<void> {
        const control = this.control(state);
        state.runAllowed = false;
        control.generation++;
        if (!control.finalTickets.size) return;
        await new Promise<void>((resolve) => control.stopWaiters.add(resolve));
    }

    async isolate(state: RunState): Promise<void> {
        state.confirmed = false;
        await this.stop(state);
    }

    beginFinal(ticket: RunTicket): void {
        this.require(ticket);
        this.control(ticket.state).finalTickets.add(ticket.id);
    }

    isFinalizing(ticket: RunTicket): boolean {
        return this.control(ticket.state).finalTickets.has(ticket.id);
    }

    finishFinal(ticket: RunTicket): void {
        const control = this.control(ticket.state);
        control.finalTickets.delete(ticket.id);
        this.release(control);
    }

    async shutdown(): Promise<void> {
        await this.suspend();
    }

    async suspend(): Promise<void> {
        this.serviceAllowed = false;
        this.serviceVersion++;
        const waits = [...this.knownControls]
            .filter((control) => control.finalTickets.size)
            .map((control) => new Promise<void>((resolve) => control.stopWaiters.add(resolve)));
        await Promise.all(waits);
    }

    resume(): void {
        this.serviceVersion++;
        this.serviceAllowed = true;
    }

    private control(state: RunState): RunControl {
        let control = this.controls.get(state);
        if (!control) {
            control = { generation: 0, finalTickets: new Set(), stopWaiters: new Set() };
            this.controls.set(state, control);
            this.knownControls.add(control);
        }
        return control;
    }

    private release(control: RunControl): void {
        if (control.finalTickets.size) return;
        for (const resolve of control.stopWaiters) resolve();
        control.stopWaiters.clear();
    }
}
